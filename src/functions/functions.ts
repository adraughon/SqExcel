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
 *    - CREATE_PLOT_CODE with options: =CREATE_PLOT_CODE(A1,"2024-01-01","2024-01-02",100,0.5,4,TRUE,TRUE,0.8,"green","normal","Temperature")
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
 * 
 * @customfunction PULL
 * @param {string[][]} sensorNames Sensor names range
 * @param {string} startDatetime Start time
 * @param {string} endDatetime End time
 * @param {string} [mode] "grid" or "points"
 * @param {string|number} [modeValue] Grid interval or point count
 * @returns {string[][]} Timestamp and sensor data
 */
export function PULL(
  sensorNames: string[][],
  startDatetime: string,
  endDatetime: string,
  mode?: string,
  modeValue?: string | number
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
    
    // Set default values for optional parameters
    const actualMode = mode || "points";
    const actualModeValue = modeValue || (actualMode === "points" ? 1000 : "15min");
    
    // Validate mode parameter
    if (actualMode !== "grid" && actualMode !== "points") {
      return [["Error: Mode must be 'grid' or 'points'"]];
    }
    
    
    // Calculate grid based on mode
    let grid: string;
    if (actualMode === "grid") {
      // Use actualModeValue as grid directly
      grid = String(actualModeValue);
      // Validate grid format
      const gridPattern = /^(\d+)(min|h|d|s)$/;
      if (!gridPattern.test(grid)) {
        return [["Error: Invalid grid format. Use format like '15min', '1h', '1d', '30s'"]];
      }
    } else {
      // actualMode === "points" - calculate grid from number of points
      const numPoints = typeof actualModeValue === 'number' ? actualModeValue : parseInt(String(actualModeValue));
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
 * NOTE: This function is DISABLED from Excel visibility because it provides poor UX.
 * Users need to know sensor names to search anyway, so it's not helpful for end users.
 * We keep this function in the codebase because it's used internally by other functions
 * to resolve sensor names to Seeq IDs when pulling data.
 * 
 * @customfunction SEARCH_SENSORS
 * @param {string[][]} sensorNames Sensor names range
 * @returns {string[][]} Search results
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
 * Gets the current value of a sensor.
 * 
 * @customfunction CURRENT
 * @param {string} sensorName Sensor name
 * @returns {any} Current value
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
 * Computes the average value of a sensor over a time range.
 * 
 * @customfunction AVERAGE
 * @param {string} sensorName Sensor name
 * @param {string} startDatetime Start time
 * @param {string} endDatetime End time
 * @returns {any} Average value
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
 * Beautiful color mapping for plots
 */
const COLOR_MAP: { [key: string]: string } = {
  'red': '#FF006E',
  'blue': '#3A86FF',
  'green': '#117D43',
  'yellow': '#FFBE0B',
  'black': '#1C1C1C',
  'gray': '#6C757D',
  'grey': '#6C757D',
  'orange': '#FF6B35',
  'purple': '#8B5CF6',
  'pink': '#F472B6',
  'brown': '#A0522D',
  'maroon': '#B03A2E',
  'cyan': '#17A2B8',
  'slate': '#2F3640'
};

/**
 * Helper function to get color hex code from color name or return hex if already provided
 */
function getColorHex(colorInput: string): string {
  // If it's already a hex color, return as-is
  if (colorInput.startsWith('#')) {
    return colorInput;
  }
  
  // Look up in color map (case insensitive)
  const normalizedColor = colorInput.toLowerCase();
  return COLOR_MAP[normalizedColor] || COLOR_MAP['green']; // Default to green if not found
}

/**
 * Creates Python plotting code with embedded sensor data.
 * 
 * @customfunction CREATE_PLOT_CODE
 * @param {string[][]} sensorNames Sensor names range
 * @param {string} startDatetime Start time
 * @param {string} endDatetime End time
 * @param {number} [points] Data points (default: 100)
 * @param {number} [height] Plot height (default: 0.3)
 * @param {number} [aspectRatio] Width/height ratio (default: 5)
 * @param {any} [showLine] Show lines (default: true)
 * @param {any} [showPoints] Show points (default: true)
 * @param {number} [opacity] Opacity 0-1 (default: 0.9)
 * @param {string} [colors] Colors comma-separated (default: green,red,blue)
 * @param {string} [style] Style: normal/minimal/sparkline (default: sparkline)
 * @param {string} [labels] Y-axis labels comma-separated (default: Sensor 1,Sensor 2)
 * @param {string} [yAxisBehavior] Y-axis: share/split (default: share)
 * @returns {string} Python code
 */
export function CREATE_PLOT_CODE(
  sensorNames: string[][],
  startDatetime: string,
  endDatetime: string,
  points?: number,
  height?: number,
  aspectRatio?: number,
  showLine?: any,
  showPoints?: any,
  opacity?: number,
  colors?: string,
  style?: string,
  labels?: string,
  yAxisBehavior?: string
): string {
  try {
    // Set default values for optional parameters
    const actualPoints = points || 100;
    const actualHeight = height || 0.3;
    const actualAspectRatio = aspectRatio || 5;
    // Handle boolean parameters with robust default handling
    // Excel may pass empty string, null, undefined, or false when parameter is omitted
    const actualShowLine = (showLine === undefined || showLine === null || showLine === "" || showLine === 0) ? true : Boolean(showLine);
    const actualShowPoints = (showPoints === undefined || showPoints === null || showPoints === "" || showPoints === 0) ? true : Boolean(showPoints);
    const actualOpacity = opacity || 0.9;
    const actualStyle = style || 'sparkline';
    const actualYAxisBehavior = yAxisBehavior || 'share';
    
    // Flatten the sensor names array and filter out empty cells
    const sensorNamesList = sensorNames
      .flat()
      .filter(name => name && name.trim() !== "");
    
    // Parse colors (comma-separated or single color) - always provide defaults using our COLOR_MAP
    const defaultColors = [
      COLOR_MAP['green'],    // #117D43
      COLOR_MAP['red'],      // #FF006E  
      COLOR_MAP['blue'],     // #3A86FF
      COLOR_MAP['purple'],   // #8B5CF6
      COLOR_MAP['orange']    // #FF6B35
    ];
    const colorList = colors ? colors.split(',').map(c => getColorHex(c.trim())) : defaultColors;
    
    // Parse labels (comma-separated or single label) - always provide defaults
    const defaultLabels = sensorNamesList.length === 1 ? ['Value'] : sensorNamesList.map((_, i) => `Sensor ${i + 1}`);
    const labelList = labels ? labels.split(',').map(l => l.trim()) : defaultLabels;
    
    if (sensorNamesList.length === 0) {
      return "Error: No sensor names provided";
    }
    
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
    const secondsPerPoint = Math.floor(timeRangeSeconds / actualPoints);
    
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
    
    // Call backend server to get sensor data for all sensors
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
    
    if (!result.success || !result.data || result.data.length === 0) {
      // Provide more detailed error information for debugging
      let errorMsg = "Error: No data returned from sensor. ";
      if (result.error) {
        errorMsg += "Error: " + result.error;
      }
      if (result.message) {
        errorMsg += " Message: " + result.message;
      }
      if (result.status_code) {
        errorMsg += " Status: " + result.status_code;
      }
      if (result.responseText) {
        errorMsg += " Response: " + result.responseText.substring(0, 200);
      }
      return errorMsg;
    }

    // Extract timestamps and sensor values for all sensors
    const timestamps: string[] = [];
    const sensorData: (number | string)[][] = [];
    const valueColumns = result.data_columns || [];
    
    if (valueColumns.length === 0) {
      return "Error: No data columns returned";
    }

    // Initialize sensor data arrays
    for (let i = 0; i < valueColumns.length; i++) {
      sensorData.push([]);
    }

    result.data.forEach((row: any) => {
      const timestamp = row.Timestamp || row.index;
      if (timestamp !== undefined) {
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
        
        // Only add timestamp once per row
        if (timestamps.length === 0 || timestamps[timestamps.length - 1] !== timestampStr) {
          timestamps.push(timestampStr);
          
          // Add sensor values for this timestamp
          valueColumns.forEach((column: string, index: number) => {
            const value = row[column];
            sensorData[index].push(value !== undefined ? value : null);
          });
        }
      }
    });

    if (timestamps.length === 0) {
      return "Error: No valid data points found";
    }

    // Generate Python code with embedded data
    const pythonCode = generatePythonPlotCode(
      timestamps,
      sensorData,
      sensorNamesList,
      colorList,
      labelList,
      actualHeight,
      actualAspectRatio,
      actualShowLine,
      actualShowPoints,
      actualOpacity,
      actualStyle,
      actualYAxisBehavior
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
  sensorData: (number | string)[][],
  sensorNames: string[],
  colors: string[],
  labels: string[] | null,
  height: number,
  aspectRatio: number,
  showLine: boolean,
  showPoints: boolean,
  opacity: number,
  style: string,
  yAxisBehavior: string
): string {
  // Convert arrays to Python list format
  const timestampsPython = timestamps.map(t => `'${t}'`).join(', ');
  
  // Convert sensor data arrays to Python format
  const sensorDataPython = sensorData.map(data => 
    '[' + data.map(v => typeof v === 'number' ? v.toString() : (v === null ? 'None' : `'${v}'`)).join(', ') + ']'
  ).join(', ');
  
  // Convert colors and labels to Python format
  const colorsPython = colors.map(c => `'${c}'`).join(', ');
  const labelsPython = labels.map(l => `'${l}'`).join(', ');

  return `import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime

def plot_sensor_data_func(timestamps, sensor_data, labels, colors, 
                         height=0.3, aspect_ratio=5, show_line=True, show_points=True, 
                         opacity=0.9, style='sparkline', y_axis_behavior='share'):
    """
    Plot sensor data with flexible styling and y-axis options.
    
    Parameters:
    - timestamps: list of timestamp strings or datetime objects
    - sensor_data: list of sensor values (single sensor) or list of lists (multiple sensors)
    - labels: list of labels for each sensor (provided by JavaScript)
    - colors: list of colors for each sensor (provided by JavaScript)
    - height: figure height
    - aspect_ratio: width/height ratio
    - show_line: whether to show connecting lines
    - show_points: whether to show data points
    - opacity: opacity of the plot elements
    - style: 'normal', 'minimal', or 'sparkline'
    - y_axis_behavior: 'share' or 'split' (only relevant for multiple sensors)
    """
    
    # Hidden parameters for fine-tuning
    dpi = 150  # Image quality (dots per inch)
    base_markersize = 2.2
    base_linewidth = 2
    max_marker_size = 4
    opacity_ratio = 0.7
    
    # Convert timestamps to datetime objects if they're strings
    if isinstance(timestamps[0], str):
        timestamps = [datetime.fromisoformat(ts.replace('Z', '+00:00')) for ts in timestamps]
    
    # Normalize sensor_data to always be a list of lists
    if not isinstance(sensor_data[0], (list, tuple)):
        sensor_data = [sensor_data]  # Single sensor case
    
    num_sensors = len(sensor_data)
    
    # Labels and colors are always provided by JavaScript
    
    # Create DataFrame
    df_data = {'timestamp': timestamps}
    for i, data in enumerate(sensor_data):
        df_data[f'sensor_{i}'] = data
    df = pd.DataFrame(df_data)
    
    # Calculate marker size and line width based on figure height
    scaled_markersize = min(base_markersize * height, max_marker_size)
    scaled_linewidth = base_linewidth / base_markersize * scaled_markersize
    
    # Set up plot styling
    linestyle = '-' if show_line else 'None'
    marker = 'o' if show_points else 'None'
    line_alpha = opacity * opacity_ratio if show_points else opacity
    
    def plot_single_sensor(ax, timestamps, values, color, label):
        """Helper function to plot a single sensor's data"""
        ax.plot(timestamps, values, linestyle=linestyle, marker=marker, 
                linewidth=scaled_linewidth, markersize=scaled_markersize, 
                color=color, alpha=line_alpha, markerfacecolor=color, 
                markeredgecolor=color, markerfacecoloralt=color, label=label)
        if show_points:
            ax.plot(timestamps, values, linestyle='None', marker='o', 
                    markersize=scaled_markersize, color=color, alpha=opacity)
    
    # Create figure with higher DPI for better quality
    plt.figure(figsize=(height * aspect_ratio, height), facecolor='#F7F7F7', dpi=dpi)
    
    if style == 'sparkline':
        # For sparkline, always overlay (ignore y_axis_behavior)
        plt.clf()
        plt.gcf().patch.set_visible(False)
        ax = plt.axes([0, 0, 1, 1])
        
        # Plot all sensors on same axis
        for i in range(num_sensors):
            plot_single_sensor(ax, df['timestamp'], df[f'sensor_{i}'], 
                             colors[i], labels[i])
        
        # Remove all spines and ticks
        for k, v in ax.spines.items():
            v.set_visible(False)
        ax.set_xticks([])
        ax.set_yticks([])
        ax.grid(False)
        ax.margins(0)
        
        # Get combined min/max values for annotations
        all_values = []
        for i in range(num_sensors):
            all_values.extend(df[f'sensor_{i}'])
        y_min = min(all_values)
        y_max = max(all_values)
        
        # Add min/max as text annotations
        ax.text(1, 0.99, f'{y_max:.1f}', transform=ax.transAxes, fontsize=8, 
                verticalalignment='top', horizontalalignment='left', color='black')
        ax.text(1, 0.01, f'{y_min:.1f}', transform=ax.transAxes, fontsize=8, 
                verticalalignment='bottom', horizontalalignment='left', color='black')
    
    else:
        # For normal and minimal styles
        ax1 = plt.gca()
        
        if num_sensors == 1 or y_axis_behavior == 'share':
            # Single sensor or shared y-axis: plot all on same axis
            for i in range(num_sensors):
                plot_single_sensor(ax1, df['timestamp'], df[f'sensor_{i}'], 
                                 colors[i], labels[i])
            
            # Style the y-axis with simple concatenated labels
            if num_sensors == 1:
                ax1.tick_params(axis='y', colors='black')
                ax1.set_ylabel(labels[0], color='black')
            else:
                ax1.tick_params(axis='y', colors='black')
                ylabel_text = ' & '.join(labels)
                ax1.set_ylabel(ylabel_text, color='black')
        
        elif num_sensors == 2 and y_axis_behavior == 'split':
            # Two sensors with split y-axis
            ax2 = ax1.twinx()
            
            # Plot sensor 1 on left axis
            plot_single_sensor(ax1, df['timestamp'], df[f'sensor_0'], 
                             colors[0], labels[0])
            
            # Plot sensor 2 on right axis
            plot_single_sensor(ax2, df['timestamp'], df[f'sensor_1'], 
                             colors[1], labels[1])
            
            # Style left y-axis to match sensor 1 color
            ax1.tick_params(axis='y', colors=colors[0])
            ax1.set_ylabel(labels[0], color=colors[0])
            
            # Style right y-axis to match sensor 2 color
            ax2.tick_params(axis='y', colors=colors[1])
            ax2.set_ylabel(labels[1], color=colors[1])
            
            # Remove grid lines from the right axis to avoid doubling
            ax2.grid(False)
        
        elif num_sensors > 2 and y_axis_behavior == 'split':
            # More than 2 sensors: fall back to shared axis
            print("Warning: Split y-axis only supported for 2 sensors. Using shared axis.")
            for i in range(num_sensors):
                plot_single_sensor(ax1, df['timestamp'], df[f'sensor_{i}'], 
                                 colors[i], labels[i])
            ax1.tick_params(axis='y', colors='black')
            ax1.set_ylabel(' & '.join(labels), color='black')
        
        # Apply style-specific formatting
        if style == 'normal':
            ax1.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d-%y %H:%M'))
            plt.xticks(rotation=45, ha='right', alpha=1.0)  # Force full opacity for axis labels
            plt.yticks(alpha=1.0)  # Force full opacity for axis labels
            # Enable grid for normal style
            ax1.grid(True, alpha=0.3)
            plt.tight_layout()
        elif style == 'minimal':
            # Only show first and last timestamp
            ax1.set_xticks([df['timestamp'].iloc[0], df['timestamp'].iloc[-1]])
            ax1.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d-%y %H:%M'))
            plt.xticks(rotation=0, ha='center', alpha=1.0)  # Force full opacity for axis labels
            plt.yticks(alpha=1.0)  # Force full opacity for axis labels
            ax1.grid(False)
            plt.subplots_adjust(left=0.1, right=0.9, top=1.0, bottom=0.15)
    
    plt.show()

# Data from Seeq sensors
timestamps = [${timestampsPython}]
sensor_data = [${sensorDataPython}]
labels = [${labelsPython}]
colors = [${colorsPython}]

# Plot parameters
height = ${height}
aspect_ratio = ${aspectRatio}
show_line = ${showLine ? 'True' : 'False'}
show_points = ${showPoints ? 'True' : 'False'}
opacity = ${opacity}
style = '${style}'
y_axis_behavior = '${yAxisBehavior}'

# Generate the plot
plot_sensor_data_func(timestamps, sensor_data, labels, colors, height, aspect_ratio, 
                     show_line, show_points, opacity, style, y_axis_behavior)`;
}

// Register custom functions with Excel
CustomFunctions.associate("PULL", PULL);
// CustomFunctions.associate("SEARCH_SENSORS", SEARCH_SENSORS); // DISABLED: Clumsy UX - users need sensor names anyway to search, not helpful for end users
CustomFunctions.associate("CURRENT", CURRENT);
CustomFunctions.associate("AVERAGE", AVERAGE);
CustomFunctions.associate("CREATE_PLOT_CODE", CREATE_PLOT_CODE);

