/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global clearInterval, console, CustomFunctions, setInterval, XMLHttpRequest */

/**
 * SqExcel Custom Functions
 * 
 * IMPORTANT: Authentication with Seeq should be done through the SqExcel taskpane, NOT through Excel functions.
 * The SEEQ_AUTH, SEEQ_AUTH_STATUS, and SEEQ_REAUTH functions are disabled and will show instructions
 * to use the taskpane instead.
 * 
 * To authenticate:
 * 1. Open the SqExcel taskpane
 * 2. Enter your Seeq server URL, access key, and password
 * 3. Click "Authenticate"
 * 4. Once authenticated, you can use SEEQ_SENSOR_DATA and other functions
 */

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(
  incrementBy: number,
  invocation: CustomFunctions.StreamingInvocation<number>
): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

// Backend server configuration
const BACKEND_URL = 'https://localhost:3000';

/**
 * Helper function to get stored Seeq credentials from the backend
 */
function getStoredCredentials(): any {
  try {
    // Get credentials from the backend server
    const result = callBackendSync('/api/seeq/credentials');
    
    if (result.success && result.credentials) {
      const credentials = result.credentials;
      // Check if credentials are still valid (not expired)
      const savedTime = new Date(credentials.timestamp);
      const now = new Date();
      const hoursDiff = (now.getTime() - savedTime.getTime()) / (1000 * 60 * 60);
      
      if (hoursDiff < 24) { // Credentials valid for 24 hours
        return credentials;
      }
    }
    return null;
  } catch (error) {
    console.log("Could not get stored credentials from backend:", error);
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
    
    console.log(`[DEBUG] Attempting to connect to: ${url}`);
    console.log(`[DEBUG] Endpoint: ${endpoint}`);
    console.log(`[DEBUG] Data:`, data);
    
    // Use synchronous request (deprecated but works in Excel custom functions)
    xhr.open(data ? 'POST' : 'GET', url, false);
    xhr.setRequestHeader('Content-Type', 'application/json');
    
    console.log(`[DEBUG] Request opened, sending...`);
    
    if (data) {
      xhr.send(JSON.stringify(data));
    } else {
      xhr.send();
    }
    
    console.log(`[DEBUG] Response received - Status: ${xhr.status}, StatusText: ${xhr.statusText}`);
    console.log(`[DEBUG] Response headers:`, xhr.getAllResponseHeaders());
    console.log(`[DEBUG] Response text:`, xhr.responseText.substring(0, 200) + '...');
    
    if (xhr.status === 200) {
      try {
        const parsed = JSON.parse(xhr.responseText);
        console.log(`[DEBUG] Successfully parsed response:`, parsed);
        return parsed;
      } catch (e) {
        console.log(`[DEBUG] Failed to parse response:`, e);
        return {
          success: false,
          error: `Failed to parse response: ${(e as Error).message}`,
          rawResponse: xhr.responseText
        };
      }
    } else {
      console.log(`[DEBUG] HTTP error: ${xhr.status} - ${xhr.statusText}`);
      return {
        success: false,
        error: `HTTP ${xhr.status}: ${xhr.statusText}`,
        responseText: xhr.responseText,
        url: url
      };
    }
  } catch (error) {
    console.log(`[DEBUG] Exception during request:`, error);
    return {
      success: false,
      error: `Backend request failed: ${error instanceof Error ? error.message : 'Unknown error'}`,
      details: error instanceof Error ? error.stack : 'No stack trace',
      url: `${BACKEND_URL}${endpoint}`
    };
  }
}

///**
// * Authenticates with a Seeq server using access key and password.
// * This function will attempt to authenticate synchronously.
// * 
// * @customfunction SEEQ_AUTH
// * @param url Seeq server URL (e.g., "https://your-server.seeq.tech")
// * @param accessKey Seeq access key
// * @param password Seeq password
// * @param authProvider Authentication provider (default: "Seeq")
// * @param ignoreSslErrors Whether to ignore SSL errors (default: false)
// * @returns Array containing authentication result
// */
/*
export function seeqAuth(
  url: string, 
  accessKey: string, 
  password: string, 
  authProvider: string = "Seeq", 
  ignoreSslErrors: boolean = false
): string[][] {
  try {
    // Validate inputs
    if (!url || !accessKey || !password) {
      return [["Error: URL, access key, and password are required"]];
    }
    
    // Call backend server
    const result = callBackendSync('/api/seeq/auth', {
      url, accessKey, password, authProvider, ignoreSslErrors
    });
    
    if (result.success) {
      return [
        ["Authentication successful"],
        ["User: " + (result.user || accessKey.substring(0, 8) + "...")],
        ["Server: " + (result.server_url || url)],
        ["Status: " + (result.message || "Authenticated")],
        ["Note: Credentials stored for future use"]
      ];
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
      
      // Check for specific network errors
      if (result.error && result.error.includes('Backend request failed')) {
        return [
          ["Network connection failed"],
          ["Error: " + result.error],
          ["Details: " + (result.details || "No additional details")],
          ["URL: " + (result.url || "Unknown")],
          ["Please check:"],
          ["1. Backend server is running on port 3000"],
          ["2. No firewall blocking localhost"],
          ["3. Excel can access localhost"]
        ];
      }
      
      return [
        ["Authentication failed"],
        ["Error: " + (result.error || result.message || "Unknown error")],
        ["Details: " + (result.details || "No additional details")],
        ["URL: " + (result.url || "Unknown")],
        ["Please check Python backend and SPy installation"]
      ];
    }
    
  } catch (error) {
    return [["Error: " + (error instanceof Error ? error.message : 'Unknown error')]];
  }
}
*/

// SEEQ_AUTH function is now disabled - use the taskpane for authentication instead
export function seeqAuth(
  url: string, 
  accessKey: string, 
  password: string, 
  authProvider: string = "Seeq", 
  ignoreSslErrors: boolean = false
): string[][] {
  return [
   // ["SEEQ_AUTH function is disabled"],
            ["Please use the SqExcel taskpane for authentication:"],
        ["1. Open the SqExcel taskpane"],
    ["2. Enter your Seeq credentials"],
    ["3. Click 'Authenticate'"],
    ["4. Then use SEEQ_SENSOR_DATA function"]
  ];
}

///**
// * Gets the current Seeq authentication status.
// * @customfunction SEEQ_AUTH_STATUS
// * @returns String indicating authentication status.
// */
/*
export function seeqAuthStatus(): string[][] {
  try {
    // Try to get stored credentials
    const credentials = getStoredCredentials();
    
    if (credentials) {
      return [
        ["Authentication Status: Credentials Available"],
        ["Server: " + credentials.url],
        ["Access Key: " + credentials.accessKey.substring(0, 8) + "..."],
        ["Saved: " + new Date(credentials.timestamp).toLocaleString()],
        ["Note: Use SEEQ_AUTH to test authentication"]
      ];
    } else {
      return [
        ["Authentication Status: No Credentials"],
        ["Message: Please use the SqExcel taskpane to authenticate first"],
        ["Note: Backend server must be running"]
      ];
    }
    
  } catch (error) {
    return [["Error: " + (error instanceof Error ? error.message : 'Unknown error')]];
  }
}
*/

// SEEQ_AUTH_STATUS function is now disabled - use the taskpane for authentication status
export function seeqAuthStatus(): string[][] {
  return [
   // ["SEEQ_AUTH_STATUS function is disabled"],
            ["Please use the SqExcel taskpane to check authentication:"],
        ["1. Open the SqExcel taskpane"],
    ["2. Check the authentication status displayed"],
    ["3. If not authenticated, enter credentials and click 'Authenticate'"]
  ];
}

/**
 * Gets the current Python/SPy authentication status.
 * @customfunction SEEQ_PYTHON_AUTH_STATUS
 * @returns String indicating Python authentication status.
 */
export function seeqPythonAuthStatus(): string[][] {
  try {
    // Call backend server to check Python authentication status
    const result = callBackendSync('/api/seeq/auth/python-status');
    
    if (result.success) {
      if (result.isAuthenticated) {
        return [
          ["Python Authentication Status: Authenticated"],
          ["User: " + (result.user || "Unknown")],
          ["Status: " + (result.message || "Success")]
        ];
      } else {
        return [
          ["Python Authentication Status: Not Authenticated"],
          ["Message: " + (result.message || "Not authenticated")]
        ];
      }
    } else {
      return [
        ["Python Authentication Status: Error"],
        ["Error: " + (result.error || result.message || "Unknown error")]
      ];
    }
    
  } catch (error) {
    return [["Error: " + (error instanceof Error ? error.message : 'Unknown error')]];
  }
}

/**
 * Re-authenticates with Seeq using stored credentials.
 * This can be used if authentication expires or fails.
 * @customfunction SEEQ_REAUTH
 * @returns String indicating re-authentication result.
 */
/*
export function seeqReauth(): string[][] {
  try {
    // Try to get stored credentials
    const credentials = getStoredCredentials();
    
    if (!credentials) {
      return [
        ["Re-authentication failed"],
        ["Error: No stored credentials"],
        ["Please use the SqExcel taskpane to authenticate first"]
      ];
    }
    
    // Call backend server to re-authenticate
    const result = callBackendSync('/api/seeq/auth', {
      url: credentials.url,
      accessKey: credentials.accessKey,
      password: credentials.password,
      authProvider: "Seeq",
      ignoreSslErrors: credentials.ignoreSsl
    });
    
    if (result.success) {
      return [
        ["Re-authentication successful"],
        ["User: " + (result.user || "Unknown")],
        ["Server: " + (result.user || "Unknown")],
        ["Status: " + (result.message || "Re-authenticated")]
      ];
    } else {
      return [
        ["Re-authentication failed"],
        ["Error: " + (result.error || result.message || "Unknown error")]
      ];
    }
    
  } catch (error) {
    return [["Error: " + (error instanceof Error ? error.message : 'Unknown error')]];
  }
}
*/

// SEEQ_REAUTH function is now disabled - use the taskpane for re-authentication
export function seeqReauth(): string[][] {
  return [
    ["SEEQ_REAUTH function is disabled"],
            ["Please use the SqExcel taskpane for re-authentication:"],
        ["1. Open the SqExcel taskpane"],
    ["2. If credentials are expired, re-enter them"],
    ["3. Click 'Authenticate' to re-authenticate"]
  ];
}

/**
 * Tests connection to a Seeq server.
 * @customfunction SEEQ_TEST_CONNECTION
 * @param url Seeq server URL
 * @returns String indicating connection status.
 */
export function seeqTestConnection(url: string): string[][] {
  try {
    // Validate input
    if (!url) {
      return [["Error: Server URL is required"]];
    }
    
    // Call backend server
    const result = callBackendSync('/api/seeq/test-connection', { url });
    
    if (result.success) {
      return [
        ["Connection Test: Successful"],
        ["Server: " + url],
        ["Status: " + (result.message || "Server is reachable")],
        ["Status Code: " + (result.status_code || "N/A")]
      ];
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
        ["Connection Test: Failed"],
        ["Server: " + url],
        ["Error: " + (result.error || result.message || "Unknown error")]
      ];
    }
    
  } catch (error) {
    return [["Error: " + (error instanceof Error ? error.message : 'Unknown error')]];
  }
}

/**
 * Gets Seeq server information.
 * @customfunction SEEQ_SERVER_INFO
 * @param url Seeq server URL
 * @returns String containing server information.
 */
export function seeqServerInfo(url: string): string[][] {
  try {
    // Validate input
    if (!url) {
      return [["Error: Server URL is required"]];
    }
    
    // Call backend server
    const result = callBackendSync('/api/seeq/server-info', { url });
    
    if (result.success) {
      const serverInfo = result.server_info;
      const infoRows: string[][] = [
        ["Server Information for: " + url],
        ["Status: " + (serverInfo.status || "Unknown")],
        ["Message: " + result.message]
      ];
      
      // Add additional server info if available
      if (serverInfo.version) {
        infoRows.push(["Version: " + serverInfo.version]);
      }
      if (serverInfo.name) {
        infoRows.push(["Name: " + serverInfo.name]);
      }
      if (serverInfo.description) {
        infoRows.push(["Description: " + serverInfo.description]);
      }
      
      return infoRows;
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
        ["Failed to get server info"],
        ["Error: " + (result.error || result.message || "Unknown error")],
        ["Server: " + url]
      ];
    }
    
  } catch (error) {
    return [["Error: " + (error instanceof Error ? error.message : 'Unknown error')]];
  }
}

/**
 * Searches for sensors in Seeq and pulls their data over a specified time range.
 * This is an array function that should be called on a range that can accommodate the output.
 * 
 * @customfunction SEEQ_SENSOR_DATA
 * @param sensorNames Range containing sensor names (e.g., B1:D1)
 * @param startDatetime Start time for data pull (ISO format: "2024-01-01T00:00:00")
 * @param endDatetime End time for data pull (ISO format: "2024-01-31T23:59:59")
 * @param grid Grid interval for data (e.g., "15min", "1h", "1d") - defaults to "15min"
 * @returns Array containing timestamp column and sensor data columns
 */
export function seeqSensorData(
  sensorNames: string[][],
  startDatetime: string,
  endDatetime: string,
  grid: string = "15min"
): string[][] {
  try {
    // Flatten the sensor names array and filter out empty cells
    const sensorNamesList = sensorNames
      .flat()
      .filter(name => name && name.trim() !== "");
    
    if (sensorNamesList.length === 0) {
      return [["Error: No sensor names provided"]];
    }
    
    // Validate datetime format
    const startDate = new Date(startDatetime);
    const endDate = new Date(endDatetime);
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return [["Error: Invalid datetime format. Use ISO format: YYYY-MM-DDTHH:MM:SS"]];
    }
    
    if (startDate >= endDate) {
      return [["Error: Start datetime must be before end datetime"]];
    }
    
    // Validate grid format
    const gridPattern = /^(\d+)(min|h|d|s)$/;
    if (!gridPattern.test(grid)) {
      return [["Error: Invalid grid format. Use format like '15min', '1h', '1d', '30s'"]];
    }

        // Check if we have stored credentials
    const authCredentials = getStoredCredentials();
    if (!authCredentials) {
      return [["Error: Not authenticated to Seeq. Please use the TSFlow taskpane to authenticate first."]];
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
      ignoreSslErrors: authCredentials.ignoreSsl
    });
    
    if (result.success && result.data && result.data.length > 0) {
      // Create header row with timestamp and sensor names
      const headers = ["Timestamp"].concat(result.data_columns || []);
      
      // Create data rows
      const dataRows = result.data.map((row: any) => {
        const timestamp = row.Timestamp || row.index || "N/A";
        const values = (result.data_columns || []).map((col: string) => {
          return row[col] !== undefined ? row[col] : "N/A";
        });
        return [timestamp].concat(values);
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
 * @customfunction SEEQ_SEARCH_SENSORS
 * @param sensorNames Range containing sensor names (e.g., B1:D1)
 * @returns Array containing search results for each sensor
 */
export function seeqSensorSearch(sensorNames: string[][]): string[][] {
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
      return [["Error: Not authenticated to Seeq. Please use SqExcel taskpane to authenticate first."]];
    }
    
    // Call backend server with credentials
    const result = callBackendSync('/api/seeq/search-sensors', {
      sensorNames: sensorNamesList,
      url: searchCredentials.url,
      accessKey: searchCredentials.accessKey,
      password: searchCredentials.password,
      authProvider: "Seeq",
      ignoreSslErrors: searchCredentials.ignoreSsl
    });
    
    if (result.success && result.search_results && result.search_results.length > 0) {
      // Create header row
      const headers = ["Sensor Name", "Seeq Name", "ID", "Type", "Status", "Path"];
      
      // Create data rows
      const dataRows = result.search_results.map((sensor: any) => {
        return [
          sensor.Original_Name || sensor.Name || "N/A",
          sensor.Name || "N/A",
          sensor.ID || "Not Found",
          sensor.Type || "N/A",
          sensor.Status || "Unknown",
          sensor.Path || "N/A"
        ];
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
        ["No search results returned"],
        ["Error: " + (result.error || result.message || "Unknown error")],
        ["Sensors: " + sensorNamesList.join(", ")]
      ];
    }
    
  } catch (error) {
    return [["Error: " + (error instanceof Error ? error.message : 'Unknown error')]];
  }
}
