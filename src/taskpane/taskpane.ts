/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Import the new Seeq API client
import { SeeqAPIClient } from './seeq-api-client';

// Version number to prove we're using the new code
const ADDIN_VERSION = "2.0.0 - Direct Seeq API Implementation";

// Global API client instance
let seeqClient: SeeqAPIClient | null = null;

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  
  // Initialize Seeq authentication
  initializeSeeqAuth();
});

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
  
  // Display version number to prove we're using new code
  console.log(`SqExcel Add-in Version: ${ADDIN_VERSION}`);
  showAuthStatus(`Add-in Version: ${ADDIN_VERSION}`, "info");
}

async function testSeeqConnection(): Promise<void> {
  const url = (document.getElementById("seeq-url") as HTMLInputElement)?.value;
  
  if (!url) {
    showAuthStatus("Please enter a Seeq server URL", "error");
    return;
  }

  showAuthStatus("Testing connection...", "loading");
  
  try {
    // Create new API client and test connection
    seeqClient = new SeeqAPIClient(url);
    
    // Test basic connection first
    const result = await seeqClient.testConnection();
    
    if (result.success) {
      showAuthStatus(result.message, "success");
      
      // Test authentication endpoint specifically
      showAuthStatus("Testing authentication endpoint...", "loading");
      const authEndpointResult = await seeqClient.testAuthEndpoint();
      
      if (authEndpointResult.success) {
        showAuthStatus(`${result.message} - ${authEndpointResult.message}`, "success");
      } else {
        showAuthStatus(`${result.message} - Warning: ${authEndpointResult.message}`, "warning");
      }
      
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
    // Use the API client to authenticate
    if (!seeqClient) {
      seeqClient = new SeeqAPIClient(url);
    }
    
    const result = await seeqClient.authenticate(accessKey, password, 'Seeq', ignoreSsl);
    
    if (result.success) {
      showAuthStatus("Authentication successful", "success");
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
    // Use the API client to search for sensors
    if (!seeqClient) {
      showAuthStatus("Please test connection first", "error");
      return;
    }
    
    const result = await seeqClient.searchSensors(sensorNames);
    
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
    // Use the API client to search and pull data
    if (!seeqClient) {
      showAuthStatus("Please test connection first", "error");
      return;
    }
    
    const result = await seeqClient.searchAndPullSensors(sensorNames, startTimeInput, endTimeInput, grid);
    
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
  
  // Clear API client state
  if (seeqClient) {
    seeqClient.logout();
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

// Load saved credentials from localStorage
function loadSavedCredentials(): void {
  const saved = localStorage.getItem("seeq_credentials");
  if (saved) {
    try {
      const credentials = JSON.parse(saved);
      
      // Populate form fields
      (document.getElementById("seeq-url") as HTMLInputElement).value = credentials.url || "";
      (document.getElementById("seeq-access-key") as HTMLInputElement).value = credentials.accessKey || "";
      (document.getElementById("seeq-password") as HTMLInputElement).value = credentials.password || "";
      (document.getElementById("ignore-ssl") as HTMLInputElement).checked = credentials.ignoreSsl || false;
      
      // Check if credentials are still valid (less than 24 hours old)
      const timestamp = new Date(credentials.timestamp);
      const now = new Date();
      const hoursDiff = (now.getTime() - timestamp.getTime()) / (1000 * 60 * 60);
      
      if (hoursDiff < 24) {
        // Credentials are still valid, update UI
        updateAuthUI(true);
        showAuthStatus("Using saved credentials", "info");
        
        // Recreate API client with saved URL
        if (credentials.url) {
          seeqClient = new SeeqAPIClient(credentials.url);
        }
      } else {
        // Credentials expired, remove them
        localStorage.removeItem("seeq_credentials");
      }
    } catch (error) {
      console.error("Failed to load saved credentials:", error);
      localStorage.removeItem("seeq_credentials");
    }
  }
}

// Update Excel function cache
function updateExcelCache(operationType: string, data: any): void {
  try {
    const jsonData = JSON.stringify(data);
    const cacheUpdate = `=SEEQ_UPDATE_CACHE("${operationType}", '${jsonData}')`;
    
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

// Display search results in the UI
function displaySearchResults(results: any[]): void {
  const resultsDiv = document.getElementById("search-results");
  if (resultsDiv && results.length > 0) {
    let html = `
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
    `;
    
    results.forEach(result => {
      html += `
        <tr>
          <td>${result.Original_Name || result.Name || "N/A"}</td>
          <td>${result.Name || "N/A"}</td>
          <td>${result.ID || "Not Found"}</td>
          <td>${result.Type || "N/A"}</td>
          <td>${result.Status || "Unknown"}</td>
        </tr>
      `;
    });
    
    html += `
          </tbody>
        </table>
      </div>
    `;
    
    resultsDiv.innerHTML = html;
  }
}

// Display data results in the UI
function displayDataResults(result: any): void {
  const resultsDiv = document.getElementById("data-results");
  if (resultsDiv && result.data && result.data.length > 0) {
    const dataToShow = result.data.slice(0, 10); // Show first 10 rows
    
    let html = `
      <h3>Data Results (${result.data.length} rows)</h3>
      <div class="results-table">
        <table>
          <thead>
            <tr>
              <th>Timestamp</th>
              ${result.data_columns.map(col => `<th>${col}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
    `;
    
    dataToShow.forEach(row => {
      html += `
        <tr>
          <td>${row.Timestamp || row.index || "N/A"}</td>
          ${result.data_columns.map(col => `<td>${row[col] !== undefined ? row[col] : "N/A"}</td>`).join('')}
        </tr>
      `;
    });
    
    html += `
          </tbody>
        </table>
        ${result.data.length > 10 ? `<p><em>Showing first 10 of ${result.data.length} rows</em></p>` : ''}
      </div>
    `;
    
    resultsDiv.innerHTML = html;
  }
}
