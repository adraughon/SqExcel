/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global clearInterval, console, CustomFunctions, setInterval, XMLHttpRequest */

/**
 * SqExcel Custom Functions
 * 
 * This add-in provides Excel functions for working with Seeq data:
 * - SEARCH_SENSORS: Search for sensors in your Seeq environment
 * - PULL: Pull time series data from Seeq sensors
 * - CREATE_PLOT_CODE: Generate Python plotting code with embedded sensor data
 * 
 * SETUP INSTRUCTIONS:
 * 1. Create a Seeq Access Key:
 *    - Go to your Seeq environment
 *    - Click on your username in the top right
 *    - Select "Create Access Key"
 *    - Copy both the Key (ID) and Password - you'll need both!
 * 
 * 2. Authenticate in Excel:
 *    - Open the SqExcel taskpane (if not visible, go to Insert > My Add-ins)
 *    - Enter your Seeq server URL (e.g., https://your-server.seeq.tech)
 *    - Enter the Access Key and Password from step 1
 *    - Click "Authenticate"
 *    - Once authenticated, you can use the Excel functions below
 * 
 * 3. Using the Functions:
 *    - SEARCH_SENSORS: =SEARCH_SENSORS(A1:C1) where A1:C1 contains sensor names
 *    - PULL: =PULL(A1:C1,"2024-01-01T00:00:00","2024-01-31T23:59:59) - defaults to 1000 points
 *    - PULL with grid: =PULL(A1:C1,"2024-01-01T00:00:00","2024-01-31T23:59:59","grid","15min")
 *    - PULL with points: =PULL(A1:C1,"2024-01-01T00:00:00","2024-01-31T23:59:59,"points",500)
 *    - CREATE_PLOT_CODE: =CREATE_PLOT_CODE(A1,"2024-01-01T00:00:00","2024-01-31T23:59:59") - basic usage
 *    - CREATE_PLOT_CODE with options: =CREATE_PLOT_CODE(A1,"2024-01-01","2024-01-02",100,0.5,4,TRUE,TRUE,0.8,"blue","normal","Temperature")
 * 
 * TROUBLESHOOTING:
 * - If you see "#NAME?" errors, make sure the add-in is properly loaded
 * - If authentication fails, check your Seeq server URL and credentials
 * - If data doesn't load, verify your sensor names exist in Seeq
 * - For detailed diagnostics, run a connection test in the taskpane
 */


// Backend server configuration
const BACKEND_URL = 'https://sqexcel.up.railway.app';

/**
 * Helper function to get stored Seeq credentials from localStorage
 */
function getStoredCredentials(): any {
  try {
    // Get credentials from localStorage (same storage used by taskpane)
    const saved = localStorage.getItem("seeq_credentials");
    if (saved) {
      const credentials = JSON.parse(saved);
      
      // Check if credentials are still valid (not expired)
      const savedTime = new Date(credentials.timestamp);
      const now = new Date();
      const hoursDiff = (now.getTime() - savedTime.getTime()) / (1000 * 60 * 60);
      
      if (hoursDiff < 24) { // Credentials valid for 24 hours
        return credentials;
      } else {
        // Credentials expired, remove them
        localStorage.removeItem("seeq_credentials");
      }
    }
    return null;
  } catch (error) {
    return null;
  }
}

/**
 * Helper function to make synchronous HTTP requests to the backend server
 * Note: This uses XMLHttpRequest synchronously, which is deprecated but works in Excel custom functions
 */
function callBackendSync(endpoint: string, data: any = null): any {
  try {
    const xhr = new XMLHttpRequest();
    const url = `${BACKEND_URL}${endpoint}`;
    
    
    // Use synchronous request (deprecated but works in Excel custom functions)
    xhr.open(data ? 'POST' : 'GET', url, false);
    xhr.setRequestHeader('Content-Type', 'application/json');
    
    
    if (data) {
      xhr.send(JSON.stringify(data));
    } else {
      xhr.send();
    }
    
    
    if (xhr.status === 200) {
      try {
        const parsed = JSON.parse(xhr.responseText);
        return parsed;
      } catch (e) {
        return {
          success: false,
          error: `Failed to parse response: ${(e as Error).message}`,
          rawResponse: xhr.responseText
        };
      }
    } else {
      return {
        success: false,
        error: `HTTP ${xhr.status}: ${xhr.statusText}`,
        responseText: xhr.responseText,
        url: url
      };
    }
  } catch (error) {
    return {
      success: false,
      error: `Backend request failed: ${error instanceof Error ? error.message : 'Unknown error'}`,
      details: error instanceof Error ? error.stack : 'No stack trace',
      url: `${BACKEND_URL}${endpoint}`
    };
  }
}


/**
 * Helper function to convert timestamps to Excel serial numbers
 * Returns Excel's internal date representation for better compatibility
 * Excel serial number = (JS timestamp / (1000 * 60 * 60 * 24)) + 25569
 * 
 * Note: These serial numbers will display as large numbers (e.g., 45870.0)
 * To see them as readable dates, users should:
 * 1. Select the timestamp column
 * 2. Right-click → Format Cells → Date
 * 3. Choose desired date format (e.g., "3/14/12 1:30 PM")
 * 
 * This function preserves the original timezone information from the backend
 * to ensure timestamps display correctly in the user's local timezone.
 */
function convertToExcelSerialNumber(timestamp: any): number {
  try {
    // Normalize the input into a Date representing the intended moment-in-time
    let date: Date | null = null;

    // If it's already an Excel serial (number or numeric string), just return it directly
    if (typeof timestamp === 'number' && timestamp > 0 && timestamp < 100000) {
      return timestamp;
    }
    if (typeof timestamp === 'string' && /^\d+(?:\.\d+)?$/.test(timestamp.trim())) {
      const asNum = parseFloat(timestamp.trim());
      if (!isNaN(asNum) && asNum > 0 && asNum < 100000) {
        return asNum;
      }
    }

    if (timestamp instanceof Date) {
      date = new Date(timestamp.getTime());
    } else if (typeof timestamp === 'number') {
      // Heuristics:
      // - Excel serials are usually < 100000
      // - Unix ms timestamps are > 10^11, Unix seconds are between 10^9 and 10^10
      if (timestamp > 1e11) {
        // treat as Unix milliseconds
        date = new Date(timestamp);
      } else if (timestamp > 1e9 && timestamp < 1e11) {
        // treat as Unix seconds
        date = new Date(timestamp * 1000);
      } else if (timestamp > 0 && timestamp < 100000) {
        // If a caller sent a number that is clearly an Excel serial, return it directly
        return timestamp;
      } else {
        // Fallback: treat as ms
        date = new Date(timestamp);
      }
    } else if (typeof timestamp === 'string') {
      // Handle ISO-like strings directly (native will keep instant with timezone if present)
      // Also handle "YYYY-MM-DD HH:MM:SS" as local time
      const localMatch = timestamp.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})$/);
      if (localMatch) {
        const [, y, m, d, hh, mm, ss] = localMatch;
        date = new Date(parseInt(y), parseInt(m) - 1, parseInt(d), parseInt(hh), parseInt(mm), parseInt(ss), 0);
      } else {
        // Try custom parser for M/D/YYYY style and AM/PM variants
        const parsed = parseDate(timestamp);
        if (parsed && !isNaN(parsed.getTime())) {
          date = parsed;
        } else {
          // Native parse (will interpret without timezone as local)
          const nd = new Date(timestamp);
          date = isNaN(nd.getTime()) ? null : nd;
        }
      }
    }

    if (!date || isNaN(date.getTime())) {
      return 0;
    }

    // Convert to Excel serial number using UTC components plus local offset once.
    // This preserves exact wall-clock minutes/seconds and avoids double-offsetting across DST.
    const tzOffsetMin = date.getTimezoneOffset();
    const utcMs = date.getTime();
    const localWallClockMs = utcMs - tzOffsetMin * 60000;
    const serial = localWallClockMs / 86400000 + 25569;

    return serial;
  } catch (_err) {
    return 0;
  }
}

/**
 * Custom date parsing function to handle various date formats consistently
 * @param dateString - Date string in various formats
 * @returns Date object or null if parsing fails
 */
function parseDate(dateString: string): Date | null {
  if (!dateString || typeof dateString !== 'string') {
    return null;
  }

  // Excel serial numbers (plain numeric strings under 100000)
  const trimmed = dateString.trim();
  if (/^\d+(?:\.\d+)?$/.test(trimmed)) {
    const serialNumber = parseFloat(trimmed);
    if (!isNaN(serialNumber) && serialNumber > 0 && serialNumber < 100000) {
      // Convert Excel serial to a JS Date that renders as LOCAL midnight.
      // 1) Compute UTC midnight for the serial
      const daysSinceUnixEpoch = serialNumber - 25569;
      const msUtc = Math.round(daysSinceUnixEpoch * 86400000);
      // 2) Apply the timezone offset at that point in time so that toString() shows local midnight
      const tzAtDateMin = new Date(msUtc).getTimezoneOffset();
      const msLocal = msUtc + tzAtDateMin * 60000;
      return new Date(msLocal);
    }
  }

  // ISO format -> rely on native
  if (dateString.includes('T')) {
    const d = new Date(dateString);
    return isNaN(d.getTime()) ? null : d;
  }

  // Handle YYYY-MM-DD HH:MM:SS as local time
  const ymdLocal = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})$/);
  if (ymdLocal) {
    const [, y, m, d, hh, mm, ss] = ymdLocal;
    return new Date(parseInt(y, 10), parseInt(m, 10) - 1, parseInt(d, 10), parseInt(hh, 10), parseInt(mm, 10), parseInt(ss, 10));
  }

  // Fallback
  const date = new Date(dateString);
  return isNaN(date.getTime()) ? null : date;
}

/**
 * Pulls time series data from Seeq sensors over a specified time range.
 * This is an array function that should be called on a range that can accommodate the output.
 * 
 * @customfunction PULL
 * @param sensorNames Range containing sensor names (e.g., B1:D1)
 * @param startDatetime Start time for data pull (ISO format: "2024-01-01T00:00:00" or "8/1/2025 0:00")
 * @param endDatetime End time for data pull (ISO format: "2024-01-31T23:59:59" or "8/3/2025 0:00")
 * @param mode Data retrieval mode: "grid" for time-based intervals or "points" for number of points - defaults to "points"
 * @param modeValue Grid interval (e.g., "15min", "1h", "1d") when mode="grid" OR number of points when mode="points" - defaults to 1000
 * @returns Array containing timestamp column (as Excel serial numbers) and sensor data columns
 * 
 * TIMEZONE BEHAVIOR:
 * - Input dates without timezone info are treated as local timezone
 * - Returned data timestamps are in the same timezone as the input (local timezone)
 * - This matches user expectations for natural date/time input
 */
export function PULL(
  sensorNames: string[][],
  startDatetime: string,
  endDatetime: string,
  mode: string = "points",
  modeValue: string | number = 1000
): string[][] {
  try {
    // Flatten the sensor names array and filter out empty cells
    const sensorNamesList = sensorNames
      .flat()
      .filter(name => name && name.trim() !== "");
    
    if (sensorNamesList.length === 0) {
      return [["Error: No sensor names provided"]];
    }
    
    // Validate datetime format and ensure consistent timezone handling
    let startDate: Date;
    let endDate: Date;
    
  // Parse dates using custom parser
  startDate = parseDate(startDatetime);
  endDate = parseDate(endDatetime);
  
  if (!startDate || !endDate) {
    return [["Error: Invalid datetime format. Use formats like: 8/1/2025 0:00 or 2024-01-01T00:00:00"]];
  }
  
  // Calculate time range in seconds (needed for both modes)
  const timeRangeMs = endDate.getTime() - startDate.getTime();
  const timeRangeSeconds = Math.floor(timeRangeMs / 1000);
    
    if (startDate >= endDate) {
      return [["Error: Start datetime must be before end datetime"]];
    }
    
    // Validate mode parameter
    if (mode !== "grid" && mode !== "points") {
      return [["Error: Mode must be 'grid' or 'points'"]];
    }
    
    
    // Calculate grid based on mode
    let grid: string;
    if (mode === "grid") {
      // Use modeValue as grid directly
      grid = String(modeValue);
      // Validate grid format
      const gridPattern = /^(\d+)(min|h|d|s)$/;
      if (!gridPattern.test(grid)) {
        return [["Error: Invalid grid format. Use format like '15min', '1h', '1d', '30s'"]];
      }
    } else {
      // mode === "points" - calculate grid from number of points
      const numPoints = typeof modeValue === 'number' ? modeValue : parseInt(String(modeValue));
      if (isNaN(numPoints) || numPoints <= 0) {
        return [["Error: Number of points must be a positive integer"]];
      }
      
      // Calculate seconds per point (must be integer)
      const secondsPerPoint = Math.floor(timeRangeSeconds / numPoints);
      
      if (secondsPerPoint < 1) {
        return [
          ["Error: Time range too short for requested number of points. Try fewer points or a longer time range."]
        ];
      }
      
      // Convert to grid format
      grid = `${secondsPerPoint}s`;
    }

        // Check if we have stored credentials
    const authCredentials = getStoredCredentials();
    if (!authCredentials) {
      return [["Error: Not authenticated to Seeq. Please use the SqExcel taskpane to authenticate first."]];
    }
    
    // Get user's timezone
    const userTimezone = (function() {
      try {
        const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
        return tz || 'UTC';
      } catch (_e) {
        return 'UTC';
      }
    })();
    
    // Call backend server with credentials
    const result = callBackendSync('/api/seeq/sensor-data', {
      sensorNames: sensorNamesList,
      startDatetime,
      endDatetime,
      grid,
      userTimezone,
      url: authCredentials.url,
      accessKey: authCredentials.accessKey,
      password: authCredentials.password,
      authProvider: "Seeq",
      ignoreSslErrors: false
    });
    
    if (result.success && result.data && result.data.length > 0) {
      // Create header row with timestamp and sensor names
      const headers = ["Timestamp"].concat(result.data_columns || []);
      
      // Create data rows with formatted timestamps
      const dataRows = result.data.map((row: any) => {
        const timestamp = row.Timestamp || row.index || "N/A";
        // Convert timestamp to Excel serial number for best compatibility
        const excelSerialTimestamp = convertToExcelSerialNumber(timestamp);
        const values = (result.data_columns || []).map((col: string) => {
          return row[col] !== undefined ? row[col] : "N/A";
        });
        return [excelSerialTimestamp].concat(values);
      });
      
      
      return [headers].concat(dataRows);
    } else {
      // Check if backend server is not running
      if (result.error && result.error.includes('Failed to fetch') || result.error.includes('NetworkError')) {
        return [
          ["Backend server not running"],
          ["Please start the backend server:"],
          ["1. Open terminal in the backend folder"],
          ["2. Run: npm install && npm start"],
          ["3. Then use this function again"]
        ];
      }
      
      return [
        ["No data returned"],
        ["Error: " + (result.error || result.message || "Unknown error")],
        ["Sensors: " + sensorNamesList.join(", ")],
        ["Time Range: " + startDatetime + " to " + endDatetime]
      ];
    }
    
  } catch (error) {
    return [["Error: " + (error instanceof Error ? error.message : 'Unknown error')]];
  }
}

/**
 * Searches for sensors in Seeq without pulling data.
 * 
 * @customfunction SEARCH_SENSORS
 * @param sensorNames Range containing sensor names (e.g., B1:D1)
 * @returns Array containing search results for each sensor
 */
export function SEARCH_SENSORS(sensorNames: string[][]): string[][] {
  try {
    // Flatten the sensor names array and filter out empty cells
    const sensorNamesList = sensorNames
      .flat()
      .filter(name => name && name.trim() !== "");
    
    if (sensorNamesList.length === 0) {
      return [["Error: No sensor names provided"]];
    }

    // Check if we have stored credentials
    const searchCredentials = getStoredCredentials();
    if (!searchCredentials) {
      return [["Error: Not authenticated to Seeq. Please use the SqExcel taskpane to authenticate first."]];
    }
    
    // Call backend server with credentials
    const result = callBackendSync('/api/seeq/search-sensors', {
      sensorNames: sensorNamesList,
      url: searchCredentials.url,
      accessKey: searchCredentials.accessKey,
      password: searchCredentials.password,
      authProvider: "Seeq",
      ignoreSslErrors: false
    });
    
    if (result.success && result.search_results && result.search_results.length > 0) {
      // Create header row
      const headers = ["Name", "ID", "Datasource Name", "Value Unit Of Measure", "Description"];
      
      // Create data rows
      const dataRows = result.search_results.map((sensor: any) => {
        return [
          sensor["Name"] || "N/A",
          sensor["ID"] || "Not Found",
          sensor["Datasource Name"] || "N/A",
          sensor["Value Unit Of Measure"] || "N/A",
          sensor["Description"] || "N/A"
        ];
      });
      
      return [headers].concat(dataRows);
    } else {
      // Check if backend server is not running
      if (result.error && result.error.includes('Failed to fetch') || result.error.includes('NetworkError')) {
        return [
          ["Backend server not running"]
        ];
      }
      
      return [
        ["No search results returned"],
        ["Error: " + (result.error || result.message || "Unknown error")],
        ["Sensors: " + sensorNamesList.join(", ")]
      ];
    }
    
  } catch (error) {
    return [["Error: " + (error instanceof Error ? error.message : 'Unknown error')]];
  }
}

/**
 * Gets the current value of a sensor at the present time.
 * Returns a single-cell scalar value.
 * 
 * @customfunction CURRENT
 * @param sensorName Name of the sensor (e.g., "Area/Tag")
 * @returns Current value as a scalar, or an error string
 */
export function CURRENT(sensorName: string): any {
  try {
    if (!sensorName || typeof sensorName !== 'string' || sensorName.trim() === '') {
      return "Error: Sensor name is required";
    }

    const authCredentials = getStoredCredentials();
    if (!authCredentials) {
      return "Error: Not authenticated to Seeq. Please use the SqExcel taskpane to authenticate first.";
    }

    // Determine a short time window ending now, using user's timezone context
    const now = new Date();
    const endDatetime = now.toISOString();
    const startDatetime = new Date(now.getTime() - 60 * 1000).toISOString(); // last 60 seconds

    // Request at 1-second grid and take the latest value
    const userTimezone = (function() {
      try {
        const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
        return tz || 'UTC';
      } catch (_e) {
        return 'UTC';
      }
    })();

    const result = callBackendSync('/api/seeq/sensor-data', {
      sensorNames: [sensorName.trim()],
      startDatetime,
      endDatetime,
      grid: '1s',
      userTimezone,
      url: authCredentials.url,
      accessKey: authCredentials.accessKey,
      password: authCredentials.password,
      authProvider: "Seeq",
      ignoreSslErrors: false
    });

    if (result.success && Array.isArray(result.data) && result.data.length > 0) {
      const columns: string[] = result.data_columns || [];
      if (!columns || columns.length === 0) {
        return "Error: No data columns returned";
      }
      const valueColumn = columns[0];
      // Take the last row's value
      const lastRow = result.data[result.data.length - 1] || {};
      const value = lastRow[valueColumn];
      return (value !== undefined && value !== null) ? value : "Error: No current value available";
    }

    return "Error: " + (result.error || result.message || "No data returned");
  } catch (error) {
    return "Error: " + (error instanceof Error ? error.message : 'Unknown error');
  }
}

/**
 * Computes the average value of a sensor over a time range using a 100-point grid.
 * Returns a single-cell scalar value.
 * 
 * @customfunction AVERAGE
 * @param sensorName Name of the sensor (e.g., "Area/Tag")
 * @param startDatetime Start time (e.g., "2024-01-01T00:00:00" or "8/1/2025 0:00")
 * @param endDatetime End time (e.g., "2024-01-31T23:59:59" or "8/1/2025 1:40")
 * @returns Average value as a scalar, or an error string
 */
export function AVERAGE(sensorName: string, startDatetime: string, endDatetime: string): any {
  try {
    if (!sensorName || typeof sensorName !== 'string' || sensorName.trim() === '') {
      return "Error: Sensor name is required";
    }

    const startDate = parseDate(startDatetime);
    const endDate = parseDate(endDatetime);

    if (!startDate || !endDate) {
      return "Error: Invalid datetime format. Use formats like: 8/1/2025 0:00 or 2024-01-01T00:00:00";
    }
    if (startDate >= endDate) {
      return "Error: Start datetime must be before end datetime";
    }

    const authCredentials = getStoredCredentials();
    if (!authCredentials) {
      return "Error: Not authenticated to Seeq. Please use the SqExcel taskpane to authenticate first.";
    }

    // Determine grid from 100 points
    const timeRangeSeconds = Math.floor((endDate.getTime() - startDate.getTime()) / 1000);
    const points = 100;
    const secondsPerPoint = Math.floor(timeRangeSeconds / points);
    if (secondsPerPoint < 1) {
      return "Error: Time range too short for 100 points. Use a longer range.";
    }
    const grid = `${secondsPerPoint}s`;

    const userTimezone = (function() {
      try {
        const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
        return tz || 'UTC';
      } catch (_e) {
        return 'UTC';
      }
    })();

    const result = callBackendSync('/api/seeq/sensor-data', {
      sensorNames: [sensorName.trim()],
      startDatetime,
      endDatetime,
      grid,
      userTimezone,
      url: authCredentials.url,
      accessKey: authCredentials.accessKey,
      password: authCredentials.password,
      authProvider: "Seeq",
      ignoreSslErrors: false
    });

    if (result.success && Array.isArray(result.data) && result.data.length > 0) {
      const columns: string[] = result.data_columns || [];
      if (!columns || columns.length === 0) {
        return "Error: No data columns returned";
      }
      const valueColumn = columns[0];
      let sum = 0;
      let count = 0;
      for (const row of result.data) {
        const v = row[valueColumn];
        if (typeof v === 'number') {
          sum += v;
          count += 1;
        } else if (typeof v === 'string') {
          const parsed = parseFloat(v);
          if (!isNaN(parsed)) {
            sum += parsed;
            count += 1;
          }
        }
      }
      if (count === 0) {
        return "Error: No numeric data to average";
      }
      return sum / count;
    }

    return "Error: " + (result.error || result.message || "No data returned");
  } catch (error) {
    return "Error: " + (error instanceof Error ? error.message : 'Unknown error');
  }
}

/**
 * Creates Python plotting code with embedded sensor data for visualization.
 * This function fetches sensor data and returns complete Python code as text.
 * 
 * @customfunction CREATE_PLOT_CODE
 * @param sensorNames Range containing sensor names (e.g., B1:D1)
 * @param startDatetime Start time for data pull (ISO format: "2024-01-01T00:00:00" or "8/1/2025 0:00")
 * @param endDatetime End time for data pull (ISO format: "2024-01-31T23:59:59" or "8/3/2025 0:00")
 * @param points Number of data points to retrieve (defaults to 100)
 * @param height Plot height in inches (defaults to 0.3)
 * @param aspectRatio Width to height ratio (defaults to 5)
 * @param showLine Whether to show connecting lines (defaults to true)
 * @param showPoints Whether to show data points (defaults to true)
 * @param opacity Point and line opacity 0-1 (defaults to 0.9)
 * @param color Plot color (defaults to 'red')
 * @param style Plot style: 'normal', 'minimal', or 'sparkline' (defaults to 'sparkline')
 * @param label Y-axis label (defaults to 'Value')
 * @returns Python code as text string
 */
export function CREATE_PLOT_CODE(
  sensorNames: string[][],
  startDatetime: string,
  endDatetime: string,
  points: number = 100,
  height: number = 0.3,
  aspectRatio: number = 5,
  showLine: boolean = true,
  showPoints: boolean = true,
  opacity: number = 0.9,
  color: string = 'red',
  style: string = 'sparkline',
  label: string = 'Value'
): string {
  try {
    // Flatten the sensor names array and filter out empty cells
    const sensorNamesList = sensorNames
      .flat()
      .filter(name => name && name.trim() !== "");
    
    if (sensorNamesList.length === 0) {
      return "Error: No sensor names provided";
    }

    // For now, we'll use the first sensor name for single sensor plotting
    const sensorName = sensorNamesList[0];
    
    // Validate datetime format
    const startDate = parseDate(startDatetime);
    const endDate = parseDate(endDatetime);
    
    if (!startDate || !endDate) {
      return "Error: Invalid datetime format. Use formats like: 8/1/2025 0:00 or 2024-01-01T00:00:00";
    }
    
    if (startDate >= endDate) {
      return "Error: Start datetime must be before end datetime";
    }

    // Calculate time range and grid for the specified number of points
    const timeRangeMs = endDate.getTime() - startDate.getTime();
    const timeRangeSeconds = Math.floor(timeRangeMs / 1000);
    const secondsPerPoint = Math.floor(timeRangeSeconds / points);
    
    if (secondsPerPoint < 1) {
      return "Error: Time range too short for requested number of points. Try fewer points or a longer time range.";
    }
    
    const grid = `${secondsPerPoint}s`;

    // Check if we have stored credentials
    const authCredentials = getStoredCredentials();
    if (!authCredentials) {
      return "Error: Not authenticated to Seeq. Please use the SqExcel taskpane to authenticate first.";
    }
    
    // Get user's timezone
    const userTimezone = (function() {
      try {
        const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
        return tz || 'UTC';
      } catch (_e) {
        return 'UTC';
      }
    })();
    
    // Call backend server to get sensor data
    const result = callBackendSync('/api/seeq/sensor-data', {
      sensorNames: [sensorName],
      startDatetime,
      endDatetime,
      grid,
      userTimezone,
      url: authCredentials.url,
      accessKey: authCredentials.accessKey,
      password: authCredentials.password,
      authProvider: "Seeq",
      ignoreSslErrors: false
    });
    
    if (!result.success || !result.data || result.data.length === 0) {
      return "Error: No data returned from sensor. " + (result.error || result.message || "Unknown error");
    }

    // Extract timestamps and sensor values
    const timestamps: string[] = [];
    const sensorValues: (number | string)[] = [];
    const valueColumn = result.data_columns?.[0];
    
    if (!valueColumn) {
      return "Error: No data columns returned";
    }

    result.data.forEach((row: any) => {
      const timestamp = row.Timestamp || row.index;
      const value = row[valueColumn];
      if (timestamp !== undefined && value !== undefined) {
        // Convert timestamp to ISO string for Python datetime parsing
        let timestampStr: string;
        if (timestamp instanceof Date) {
          timestampStr = timestamp.toISOString();
        } else if (typeof timestamp === 'number') {
          // Handle Excel serial numbers or Unix timestamps
          const date = new Date(timestamp > 100000 ? timestamp : (timestamp - 25569) * 86400000);
          timestampStr = date.toISOString();
        } else {
          timestampStr = String(timestamp);
        }
        timestamps.push(timestampStr);
        sensorValues.push(value);
      }
    });

    if (timestamps.length === 0) {
      return "Error: No valid data points found";
    }

    // Generate Python code with embedded data
    const pythonCode = generatePythonPlotCode(
      timestamps,
      sensorValues,
      height,
      aspectRatio,
      showLine,
      showPoints,
      opacity,
      color,
      style,
      label
    );

    return pythonCode;
    
  } catch (error) {
    return "Error: " + (error instanceof Error ? error.message : 'Unknown error');
  }
}

/**
 * Helper function to generate Python plotting code with embedded data
 */
function generatePythonPlotCode(
  timestamps: string[],
  sensorValues: (number | string)[],
  height: number,
  aspectRatio: number,
  showLine: boolean,
  showPoints: boolean,
  opacity: number,
  color: string,
  style: string,
  label: string
): string {
  // Convert arrays to Python list format
  const timestampsPython = timestamps.map(t => `'${t}'`).join(', ');
  const valuesPython = sensorValues.map(v => 
    typeof v === 'number' ? v.toString() : `'${v}'`
  ).join(', ');

  return `import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime

# Data from Seeq sensor
timestamps = [${timestampsPython}]
sensor_values = [${valuesPython}]

# Convert timestamps to datetime objects
timestamps = [datetime.fromisoformat(ts.replace('Z', '+00:00')) for ts in timestamps]

# Create DataFrame
df = pd.DataFrame({'timestamp': timestamps, 'sensor_value': sensor_values})

# Plot parameters
height = ${height}
aspect_ratio = ${aspectRatio}
show_line = ${showLine ? 'True' : 'False'}
show_points = ${showPoints ? 'True' : 'False'}
opacity = ${opacity}
color = '${color}'
style = '${style}'  # 'normal', 'minimal', or 'sparkline'
label = '${label}'

# Calculate marker size and line width based on figure height
base_markersize = 2.2
base_linewidth = 2
max_marker_size = 4
opacity_ratio = 0.7  # lower means fainter lines

scaled_markersize = min(base_markersize * height, max_marker_size)
scaled_linewidth = base_linewidth / base_markersize * scaled_markersize

plt.figure(figsize=(height * aspect_ratio, height), facecolor='#F7F7F7')
linestyle = '-' if show_line else 'None'
marker = 'o' if show_points else 'None'
line_alpha = opacity * opacity_ratio if show_points else opacity
plt.plot(df['timestamp'], df['sensor_value'], linestyle=linestyle, marker=marker, linewidth=scaled_linewidth, markersize=scaled_markersize, color=color, alpha=line_alpha, markerfacecolor=color, markeredgecolor=color, markerfacecoloralt=color)
if show_points:
    plt.plot(df['timestamp'], df['sensor_value'], linestyle='None', marker='o', markersize=scaled_markersize, color=color, alpha=opacity)

if style == 'normal':
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%m-%d-%y %H:%M'))
    plt.xticks(rotation=45, ha='right')
    plt.ylabel(label)
    plt.tight_layout()
elif style == 'minimal':
    # Only show first and last timestamp
    ax = plt.gca()
    ax.set_xticks([df['timestamp'].iloc[0], df['timestamp'].iloc[-1]])
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d-%y %H:%M'))
    plt.xticks(rotation=0, ha='center')  # Flat text direction
    plt.ylabel(label)
    # Remove grid lines
    ax.grid(False)
    # Maximize figure area - remove margins
    plt.subplots_adjust(left=0.1, right=1.0, top=1.0, bottom=0.15)
elif style == 'sparkline':
    # Create a new figure without frame for sparkline
    plt.clf()  # Clear the figure
    plt.gcf().patch.set_visible(False)  # Remove figure background/frame
    ax = plt.axes([0, 0, 1, 1])  # Create axes that fill entire figure
    
    # Plot the data again since we cleared the figure
    plt.plot(df['timestamp'], df['sensor_value'], linestyle=linestyle, marker=marker, 
             linewidth=scaled_linewidth, markersize=scaled_markersize, color=color, 
             alpha=line_alpha, markerfacecolor=color, markeredgecolor=color)
    if show_points:
        plt.plot(df['timestamp'], df['sensor_value'], linestyle='None', marker='o', 
                 markersize=scaled_markersize, color=color, alpha=opacity)
    
    # Remove all spines
    for k, v in ax.spines.items():
        v.set_visible(False)
    ax.set_xticks([])
    ax.set_yticks([])
    
    # Remove grid and set margins to zero
    ax.grid(False)
    ax.margins(0)
    
    # Get min/max values for annotations
    y_min = df['sensor_value'].min()
    y_max = df['sensor_value'].max()
    
    # Add min/max as text annotations positioned within plot area
    ax.text(1, 0.99, f'{y_max:.1f}', transform=ax.transAxes, fontsize=8, 
            verticalalignment='top', horizontalalignment='left', color='black')
    ax.text(1, 0.01, f'{y_min:.1f}', transform=ax.transAxes, fontsize=8, 
            verticalalignment='bottom', horizontalalignment='left', color='black')

plt.show()`;
}
