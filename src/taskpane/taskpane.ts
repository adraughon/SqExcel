/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Backend server configuration
const BACKEND_URL = 'https://localhost:3000';

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  
  // Initialize Seeq authentication
  initializeSeeqAuth();
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

// Seeq Authentication Interface
interface SeeqAuthResult {
  success: boolean;
  message: string;
  user?: string;
  server_url?: string;
  error?: string;
}

interface SeeqConnectionResult {
  success: boolean;
  message: string;
  status_code?: number;
  error?: string;
}

// Seeq Authentication Functions
function initializeSeeqAuth(): void {
  const testConnectionBtn = document.getElementById("test-connection") as HTMLButtonElement;
  const authenticateBtn = document.getElementById("authenticate") as HTMLButtonElement;
  const logoutBtn = document.getElementById("logout") as HTMLButtonElement;
  const searchSensorsBtn = document.getElementById("search-sensors") as HTMLButtonElement;
  const pullDataBtn = document.getElementById("pull-data") as HTMLButtonElement;

  if (testConnectionBtn) {
    testConnectionBtn.onclick = testSeeqConnection;
  }
  
  if (authenticateBtn) {
    authenticateBtn.onclick = authenticateWithSeeq;
  }
  
  if (logoutBtn) {
    logoutBtn.onclick = logoutFromSeeq;
  }

  if (searchSensorsBtn) {
    searchSensorsBtn.onclick = searchSensors;
  }

  if (pullDataBtn) {
    pullDataBtn.onclick = pullSensorData;
  }

  // Load saved credentials if available
  loadSavedCredentials();
}

async function testSeeqConnection(): Promise<void> {
  const url = (document.getElementById("seeq-url") as HTMLInputElement)?.value;
  
  if (!url) {
    showAuthStatus("Please enter a Seeq server URL", "error");
    return;
  }

  showAuthStatus("Testing connection...", "loading");
  
  try {
    // Call the Python backend to test connection
    const result = await testConnection(url);
    
    if (result.success) {
      showAuthStatus(result.message, "success");
      // Update Excel function cache
      updateExcelCache("auth", result);
    } else {
      showAuthStatus(result.message, "error");
    }
  } catch (error) {
    showAuthStatus(`Connection test failed: ${error}`, "error");
  }
}

async function authenticateWithSeeq(): Promise<void> {
  const url = (document.getElementById("seeq-url") as HTMLInputElement)?.value;
  const accessKey = (document.getElementById("seeq-access-key") as HTMLInputElement)?.value;
  const password = (document.getElementById("seeq-password") as HTMLInputElement)?.value;
  const ignoreSsl = (document.getElementById("ignore-ssl") as HTMLInputElement)?.checked;

  if (!url || !accessKey || !password) {
    showAuthStatus("Please fill in all required fields", "error");
    return;
  }

  showAuthStatus("Authenticating...", "loading");
  
  try {
    // Call the Python backend to authenticate
    const result = await authenticate(url, accessKey, password, ignoreSsl);
    
    if (result.success) {
      showAuthStatus(result.message, "success");
      saveCredentials(url, accessKey, password, ignoreSsl);
      updateAuthUI(true);
      // Update Excel function cache
      updateExcelCache("auth", result);
    } else {
      showAuthStatus(result.message, "error");
    }
  } catch (error) {
    showAuthStatus(`Authentication failed: ${error}`, "error");
  }
}

async function searchSensors(): Promise<void> {
  const sensorNamesInput = (document.getElementById("sensor-names") as HTMLInputElement)?.value;
  
  if (!sensorNamesInput) {
    showAuthStatus("Please enter sensor names to search for", "error");
    return;
  }

  const sensorNames = sensorNamesInput.split(',').map(name => name.trim()).filter(name => name);
  
  if (sensorNames.length === 0) {
    showAuthStatus("Please enter valid sensor names", "error");
    return;
  }

  showAuthStatus("Searching for sensors...", "loading");
  
  try {
    // Call the Python backend to search for sensors
    const result = await searchSensorsOnly(sensorNames);
    
    if (result.success) {
      showAuthStatus(`Found ${result.sensor_count} sensors`, "success");
      // Update Excel function cache
      updateExcelCache("search", result);
      // Display results
      displaySearchResults(result.search_results);
    } else {
      showAuthStatus(result.message, "error");
    }
  } catch (error) {
    showAuthStatus(`Sensor search failed: ${error}`, "error");
  }
}

async function pullSensorData(): Promise<void> {
  const sensorNamesInput = (document.getElementById("sensor-names") as HTMLInputElement)?.value;
  const startTimeInput = (document.getElementById("start-time") as HTMLInputElement)?.value;
  const endTimeInput = (document.getElementById("end-time") as HTMLInputElement)?.value;
  const gridInput = (document.getElementById("grid-interval") as HTMLInputElement)?.value;
  
  if (!sensorNamesInput || !startTimeInput || !endTimeInput) {
    showAuthStatus("Please fill in all required fields", "error");
    return;
  }

  const sensorNames = sensorNamesInput.split(',').map(name => name.trim()).filter(name => name);
  const grid = gridInput || "15min";
  
  if (sensorNames.length === 0) {
    showAuthStatus("Please enter valid sensor names", "error");
    return;
  }

  showAuthStatus("Pulling sensor data...", "loading");
  
  try {
    // Call the Python backend to search and pull data
    const result = await searchAndPullSensors(sensorNames, startTimeInput, endTimeInput, grid);
    
    if (result.success) {
      showAuthStatus(`Retrieved data for ${result.sensor_count} sensors`, "success");
      // Update Excel function cache
      updateExcelCache("data", result);
      // Display results
      displayDataResults(result);
    } else {
      showAuthStatus(result.message, "error");
    }
  } catch (error) {
    showAuthStatus(`Data retrieval failed: ${error}`, "error");
  }
}

function logoutFromSeeq(): void {
  // Clear saved credentials
  localStorage.removeItem("seeq_credentials");
  
  // Clear credentials from backend so Excel functions won't work
  try {
    fetch(`${BACKEND_URL}/api/seeq/credentials`, {
      method: 'DELETE'
    }).catch(error => {
      console.log("Could not clear backend credentials:", error);
    });
  } catch (error) {
    console.log("Could not clear backend credentials:", error);
  }
  
  // Update UI
  updateAuthUI(false);
  showAuthStatus("Logged out successfully", "info");
  
  // Clear form fields
  (document.getElementById("seeq-url") as HTMLInputElement).value = "";
  (document.getElementById("seeq-access-key") as HTMLInputElement).value = "";
  (document.getElementById("seeq-password") as HTMLInputElement).value = "";
  (document.getElementById("ignore-ssl") as HTMLInputElement).checked = false;
}

function showAuthStatus(message: string, type: "success" | "error" | "info" | "loading"): void {
  const statusElement = document.getElementById("auth-status");
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.className = `auth-status ${type}`;
  }
}

function updateAuthUI(isAuthenticated: boolean): void {
  const testBtn = document.getElementById("test-connection") as HTMLButtonElement;
  const authBtn = document.getElementById("authenticate") as HTMLButtonElement;
  const logoutBtn = document.getElementById("logout") as HTMLButtonElement;
  const urlInput = document.getElementById("seeq-url") as HTMLInputElement;
  const accessKeyInput = document.getElementById("seeq-access-key") as HTMLInputElement;
  const passwordInput = document.getElementById("seeq-password") as HTMLInputElement;
  const sslCheckbox = document.getElementById("ignore-ssl") as HTMLInputElement;

  if (isAuthenticated) {
    testBtn.style.display = "none";
    authBtn.style.display = "none";
    logoutBtn.style.display = "inline-block";
    
    // Disable form inputs
    urlInput.disabled = true;
    accessKeyInput.disabled = true;
    passwordInput.disabled = true;
    sslCheckbox.disabled = true;
  } else {
    testBtn.style.display = "inline-block";
    authBtn.style.display = "inline-block";
    logoutBtn.style.display = "none";
    
    // Enable form inputs
    urlInput.disabled = false;
    accessKeyInput.disabled = false;
    passwordInput.disabled = false;
    sslCheckbox.disabled = false;
  }
}

function saveCredentials(url: string, accessKey: string, password: string, ignoreSsl: boolean): void {
  const credentials = {
    url,
    accessKey,
    password,
    ignoreSsl,
    timestamp: new Date().toISOString()
  };
  
  localStorage.setItem("seeq_credentials", JSON.stringify(credentials));
}

// Function to update credentials in the backend for Excel functions
async function updateBackendCredentials(url: string, accessKey: string, password: string, ignoreSsl: boolean): Promise<void> {
  try {
    const response = await fetch(`${BACKEND_URL}/api/seeq/credentials`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ 
        url, 
        accessKey, 
        password, 
        authProvider: 'Seeq', 
        ignoreSslErrors: ignoreSsl,
        timestamp: new Date().toISOString()
      })
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    if (result.success) {
      showAuthStatus("Credentials updated for Excel functions", "success");
    } else {
      showAuthStatus("Failed to update credentials for Excel functions", "error");
    }
  } catch (error) {
    showAuthStatus(`Failed to update backend credentials: ${error}`, "error");
  }
}

function loadSavedCredentials(): void {
  const saved = localStorage.getItem("seeq_credentials");
  if (saved) {
    try {
      const credentials = JSON.parse(saved);
      
      (document.getElementById("seeq-url") as HTMLInputElement).value = credentials.url || "";
      (document.getElementById("seeq-access-key") as HTMLInputElement).value = credentials.accessKey || "";
      (document.getElementById("seeq-password") as HTMLInputElement).value = credentials.password || "";
      (document.getElementById("ignore-ssl") as HTMLInputElement).checked = credentials.ignoreSsl || false;
      
      // Check if credentials are still valid (not expired)
      const savedTime = new Date(credentials.timestamp);
      const now = new Date();
      const hoursDiff = (now.getTime() - savedTime.getTime()) / (1000 * 60 * 60);
      
      if (hoursDiff < 24) { // Credentials valid for 24 hours
        updateAuthUI(true);
        showAuthStatus("Using saved credentials", "info");
      } else {
        // Clear expired credentials
        localStorage.removeItem("seeq_credentials");
      }
    } catch (error) {
      console.error("Failed to load saved credentials:", error);
    }
  }
}



// Python backend integration functions
async function testConnection(url: string): Promise<SeeqConnectionResult> {
  try {
    const response = await fetch(`${BACKEND_URL}/api/seeq/test-connection`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ url })
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    return await response.json();
  } catch (error) {
    // Fallback to direct Python call if API not available
    return await callPythonBackend('test_connection', [url]);
  }
}

async function authenticate(url: string, accessKey: string, password: string, ignoreSsl: boolean): Promise<SeeqAuthResult> {
  try {
    const response = await fetch(`${BACKEND_URL}/api/seeq/auth`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ url, accessKey, password, authProvider: 'Seeq', ignoreSslErrors: ignoreSsl })
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    
    // If authentication is successful, also send credentials to the credentials endpoint
    // so Excel functions can access them
    if (result.success) {
      try {
        await fetch(`${BACKEND_URL}/api/seeq/credentials`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ 
            url, 
            accessKey, 
            password, 
            authProvider: 'Seeq', 
            ignoreSslErrors: ignoreSsl,
            timestamp: new Date().toISOString()
          })
        });
      } catch (credError) {
        console.log("Could not update credentials endpoint:", credError);
        // Don't fail authentication if credentials endpoint update fails
      }
    }
    
    return result;
  } catch (error) {
    // Fallback to direct Python call if API not available
    return await callPythonBackend('authenticate_seeq', [url, accessKey, password, 'Seeq', ignoreSsl]);
  }
}

async function searchSensorsOnly(sensorNames: string[]): Promise<any> {
  try {
    // Get stored credentials
    const saved = localStorage.getItem("seeq_credentials");
    if (!saved) {
      throw new Error('No stored credentials. Please authenticate first.');
    }
    
    const credentials = JSON.parse(saved);
    
    const response = await fetch(`${BACKEND_URL}/api/seeq/search-sensors`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ 
        sensorNames,
        url: credentials.url,
        accessKey: credentials.accessKey,
        password: credentials.password,
        authProvider: 'Seeq',
        ignoreSslErrors: credentials.ignoreSsl
      })
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    return await response.json();
  } catch (error) {
    // Fallback to direct Python call if API not available
    return await callPythonBackend('search_sensors_only', [sensorNames]);
  }
}

async function searchAndPullSensors(sensorNames: string[], startTime: string, endTime: string, grid: string): Promise<any> {
  try {
    // Get stored credentials
    const saved = localStorage.getItem("seeq_credentials");
    if (!saved) {
      throw new Error('No stored credentials. Please authenticate first.');
    }
    
    const credentials = JSON.parse(saved);
    
    const response = await fetch(`${BACKEND_URL}/api/seeq/sensor-data`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ 
        sensorNames, 
        startDatetime: startTime, 
        endDatetime: endTime, 
        grid,
        url: credentials.url,
        accessKey: credentials.accessKey,
        password: credentials.password,
        authProvider: 'Seeq',
        ignoreSslErrors: credentials.ignoreSsl
      })
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    return await response.json();
  } catch (error) {
    // Fallback to direct Python call if API not available
    return await callPythonBackend('search_and_pull_sensors', [sensorNames, startTime, endTime, grid]);
  }
}

// Direct Python backend call (fallback)
async function callPythonBackend(functionName: string, args: any[]): Promise<any> {
  // This would be implemented to directly call the Python backend
  // For now, return a placeholder
  return {
    success: false,
    message: "Direct Python backend calls not yet implemented",
    error: "Use taskpane interface instead"
  };
}

// Update Excel function cache
function updateExcelCache(operationType: string, data: any): void {
  try {
    // Call the Excel function to update cache
    const dataJson = JSON.stringify(data);
    const cacheUpdate = `=SEEQ_UPDATE_CACHE("${operationType}", '${dataJson}')`;
    
    // Insert the cache update into the active worksheet
    Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange("A1");
      range.values = [[cacheUpdate]];
      await context.sync();
    }).catch(error => {
      console.log("Could not update Excel cache automatically:", error);
      console.log("Manual cache update required:", cacheUpdate);
    });
    
  } catch (error) {
    console.error("Failed to update Excel cache:", error);
  }
}

// Display functions for results
function displaySearchResults(results: any[]): void {
  const resultsDiv = document.getElementById("search-results");
  if (resultsDiv) {
    resultsDiv.innerHTML = `
      <h3>Search Results (${results.length} sensors)</h3>
      <div class="results-table">
        <table>
          <thead>
            <tr>
              <th>Original Name</th>
              <th>Seeq Name</th>
              <th>ID</th>
              <th>Type</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${results.map(sensor => `
              <tr>
                <td>${sensor.Original_Name || sensor.Name || 'N/A'}</td>
                <td>${sensor.Name || 'N/A'}</td>
                <td>${sensor.ID || 'Not Found'}</td>
                <td>${sensor.Type || 'N/A'}</td>
                <td>${sensor.Status || 'Unknown'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }
}

function displayDataResults(result: any): void {
  const resultsDiv = document.getElementById("data-results");
  if (resultsDiv && result.data && result.data.length > 0) {
    const sampleData = result.data.slice(0, 10); // Show first 10 rows
    resultsDiv.innerHTML = `
      <h3>Data Results (${result.data.length} rows)</h3>
      <div class="results-table">
        <table>
          <thead>
            <tr>
              <th>Timestamp</th>
              ${result.data_columns.map((col: string) => `<th>${col}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
            ${sampleData.map((row: any) => `
              <tr>
                <td>${row.Timestamp || row.index || 'N/A'}</td>
                ${result.data_columns.map((col: string) => `<td>${row[col] !== undefined ? row[col] : 'N/A'}</td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
        ${result.data.length > 10 ? `<p><em>Showing first 10 of ${result.data.length} rows</em></p>` : ''}
      </div>
    `;
  }
}
