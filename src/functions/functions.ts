/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global clearInterval, console, CustomFunctions, setInterval, XMLHttpRequest */

/**
 * SqExcel Custom Functions
 * 
 * This add-in provides 2 Excel functions for working with Seeq data:
 * - SEARCH_SENSORS: Search for sensors in your Seeq environment
 * - PULL: Pull time series data from Seeq sensors
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
 */
function convertToExcelSerialNumber(timestamp: any): number {
  try {
    // Handle different timestamp formats that might come from the backend
    let date: Date;
    
    if (typeof timestamp === 'string') {
      // Try to parse ISO string or other date formats
      date = new Date(timestamp);
    } else if (timestamp instanceof Date) {
      date = timestamp;
    } else if (typeof timestamp === 'number') {
      // Handle Unix timestamp
      date = new Date(timestamp);
    } else {
      // Fallback for unknown formats
      return 0; // Return 0 for invalid dates
    }
    
    // Check if date is valid
    if (isNaN(date.getTime())) {
      return 0; // Return 0 for invalid dates
    }
    
    // Convert to Excel serial number
    // Excel serial number = (JS timestamp / (1000 * 60 * 60 * 24)) + 25569
    // Where 25569 is the number of days between 1900-01-01 and 1970-01-01
    const excelSerial = (date.getTime() / (1000 * 60 * 60 * 24)) + 25569;
    
    return excelSerial;
  } catch (error) {
    // If any error occurs during conversion, return 0
    return 0;
  }
}

/**
 * Pulls time series data from Seeq sensors over a specified time range.
 * This is an array function that should be called on a range that can accommodate the output.
 * 
 * @customfunction PULL
 * @param sensorNames Range containing sensor names (e.g., B1:D1)
 * @param startDatetime Start time for data pull (ISO format: "2024-01-01T00:00:00")
 * @param endDatetime End time for data pull (ISO format: "2024-01-31T23:59:59")
 * @param mode Data retrieval mode: "grid" for time-based intervals or "points" for number of points - defaults to "points"
 * @param modeValue Grid interval (e.g., "15min", "1h", "1d") when mode="grid" OR number of points when mode="points" - defaults to 1000
 * @returns Array containing timestamp column (as Excel serial numbers) and sensor data columns
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
    
    // Try to parse dates consistently
    try {
      // If the input looks like a local date format, parse it as local time
      if (startDatetime.includes('/') && !startDatetime.includes('T')) {
        // Format like "8/1/2025 0:00" - treat as local time
        startDate = new Date(startDatetime);
        endDate = new Date(endDatetime);
      } else {
        // ISO format or other - use as is
        startDate = new Date(startDatetime);
        endDate = new Date(endDatetime);
      }
    } catch (e) {
      return [["Error: Invalid datetime format. Use ISO format: YYYY-MM-DDTHH:MM:SS"]];
    }
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return [["Error: Invalid datetime format. Use ISO format: YYYY-MM-DDTHH:MM:SS"]];
    }
    
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
      
    // Calculate time range in seconds
    const timeRangeMs = endDate.getTime() - startDate.getTime();
    const timeRangeSeconds = Math.floor(timeRangeMs / 1000);
      
      // Calculate seconds per point (must be integer)
      const secondsPerPoint = Math.floor(timeRangeSeconds / numPoints);
      
      if (secondsPerPoint < 1) {
        return [
          ["Error: Time range too short for requested number of points. Try fewer points or a longer time range."],
          [`Debug: Time range: ${timeRangeSeconds}s (${(timeRangeSeconds/3600).toFixed(2)}h), Points: ${numPoints}`],
          [`Debug: Start: ${startDate.toISOString()}, End: ${endDate.toISOString()}`]
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
    
    // Call backend server with credentials
    const result = callBackendSync('/api/seeq/sensor-data', {
      sensorNames: sensorNamesList,
      startDatetime,
      endDatetime,
      grid,
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
      
      // Add debug info as the first cell of the first data row
      if (dataRows.length > 0) {
        dataRows[0][0] = `Debug: Grid=${grid}, Points=${modeValue}, Actual rows=${dataRows.length}`;
      }
      
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
