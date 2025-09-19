/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Import CSS for styling
import './taskpane.css';

// Import the new Seeq API client
import { SeeqAPIClient } from './seeq-api-client';

// Import color map from functions
import { COLOR_MAP } from '../functions/functions';

// Version will be dynamically loaded from version.json
let ADDIN_VERSION = "Loading...";

// Global API client instance
let seeqClient: SeeqAPIClient | null = null;

// Load version information from version.json
async function loadVersionInfo(): Promise<void> {
  try {
    // Try relative path first, then fallback to full URL
    let response = await fetch('./version.json');
    
    if (!response.ok) {
      // Fallback to full URL if relative path fails
      response = await fetch('https://adraughon.github.io/SqExcel/version.json');
    }
    
    if (response.ok) {
      const versionData = await response.json();
      ADDIN_VERSION = versionData.version;
    } else {
      ADDIN_VERSION = "Version info unavailable";
    }
  } catch (error) {
    ADDIN_VERSION = "Version info unavailable";
  }
}

// The initialize function must be run each time a new page is loaded
Office.onReady(async () => {
  const sideloadMsg = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");
  if (sideloadMsg) sideloadMsg.style.display = "none";
  if (appBody) appBody.style.display = "flex";
  
  // Load version information first
  await loadVersionInfo();
  
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
  const authenticateBtn = document.getElementById("authenticate") as HTMLButtonElement;
  const logoutBtn = document.getElementById("logout") as HTMLButtonElement;
  const searchSensorsBtn = document.getElementById("search-sensors") as HTMLButtonElement;
  const pullDataBtn = document.getElementById("pull-data") as HTMLButtonElement;

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

  // Add event listener for mode selection change
  const dataModeSelect = document.getElementById("data-mode") as HTMLSelectElement;
  if (dataModeSelect) {
    dataModeSelect.onchange = handleModeChange;
  }

  // Load saved credentials if available
  loadSavedCredentials();
  
  // Display version number in header
  const versionDisplay = document.getElementById("version-display");
  if (versionDisplay) {
    versionDisplay.textContent = `v${ADDIN_VERSION}`;
  }
  
  // Generate dynamic color map
  generateColorMap();
}

function generateColorMap(): void {
  // Find the color grid container in the HTML
  const colorGrid = document.querySelector('.grid.grid-cols-2.md\\:grid-cols-3.gap-2.text-xs');
  
  if (colorGrid) {
    // Clear existing content
    colorGrid.innerHTML = '';
    
    // Generate color swatches from COLOR_MAP
    Object.entries(COLOR_MAP).forEach(([colorName, colorHex]) => {
      // Skip duplicate gray/grey entries
      if (colorName === 'grey') return;
      
      const colorDiv = document.createElement('div');
      colorDiv.className = 'flex items-center gap-2';
      colorDiv.innerHTML = `
        <div class="w-3 h-3 rounded" style="background-color: ${colorHex};"></div>
        <span class="text-gray-700">"${colorName}"</span>
      `;
      colorGrid.appendChild(colorDiv);
    });
  }
}

function handleModeChange(): void {
  const dataModeSelect = document.getElementById("data-mode") as HTMLSelectElement;
  const modeValueInput = document.getElementById("mode-value") as HTMLInputElement;
  const modeValueLabel = document.getElementById("mode-value-label") as HTMLLabelElement;
  
  if (!dataModeSelect || !modeValueInput || !modeValueLabel) return;
  
  const selectedMode = dataModeSelect.value;
  
  if (selectedMode === "grid") {
    modeValueLabel.textContent = "Grid Interval:";
    modeValueInput.placeholder = "15min";
    modeValueInput.value = "15min";
    modeValueInput.type = "text";
  } else {
    modeValueLabel.textContent = "Number of Points:";
    modeValueInput.placeholder = "1000";
    modeValueInput.value = "1000";
    modeValueInput.type = "number";
  }
}

async function authenticateWithSeeq(): Promise<void> {
  const url = (document.getElementById("seeq-url") as HTMLInputElement)?.value;
  const accessKey = (document.getElementById("seeq-access-key") as HTMLInputElement)?.value;
  const password = (document.getElementById("seeq-password") as HTMLInputElement)?.value;

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
    
    const result = await seeqClient.authenticate(accessKey, password, 'Seeq', false);
    
    if (result.success) {
      showAuthStatus("Authentication successful", "success");
      saveCredentials(url, accessKey, password, false);
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
  const dataMode = (document.getElementById("data-mode") as HTMLSelectElement)?.value;
  const modeValueInput = (document.getElementById("mode-value") as HTMLInputElement)?.value;
  
  if (!sensorNamesInput || !startTimeInput || !endTimeInput) {
    showAuthStatus("Please fill in all required fields", "error");
    return;
  }

  const sensorNames = sensorNamesInput.split(',').map(name => name.trim()).filter(name => name);
  const mode = dataMode || "points";
  const modeValue = modeValueInput || (mode === "points" ? "1000" : "15min");
  
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
    
    const result = await seeqClient.searchAndPullSensors(sensorNames, startTimeInput, endTimeInput, mode, modeValue);
    
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
}

function showAuthStatus(message: string, type: "success" | "error" | "info" | "loading" | "warning"): void {
  const statusElement = document.getElementById("auth-status");
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.className = `auth-status ${type}`;
  }
}

function updateAuthUI(isAuthenticated: boolean): void {
  const authBtn = document.getElementById("authenticate") as HTMLButtonElement;
  const logoutBtn = document.getElementById("logout") as HTMLButtonElement;
  const urlInput = document.getElementById("seeq-url") as HTMLInputElement;
  const accessKeyInput = document.getElementById("seeq-access-key") as HTMLInputElement;
  const passwordInput = document.getElementById("seeq-password") as HTMLInputElement;

  if (isAuthenticated) {
    authBtn.style.display = "none";
    logoutBtn.style.display = "inline-block";
    
    // Disable form inputs
    urlInput.disabled = true;
    accessKeyInput.disabled = true;
    passwordInput.disabled = true;
  } else {
    authBtn.style.display = "inline-block";
    logoutBtn.style.display = "none";
    
    // Enable form inputs
    urlInput.disabled = false;
    accessKeyInput.disabled = false;
    passwordInput.disabled = false;
  }
}

function saveCredentials(url: string, accessKey: string, password: string, ignoreSsl: boolean): void {
  const credentials = {
    url,
    accessKey,
    password,
    ignoreSsl: false,
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
      
      // Check if credentials are still valid (less than 24 hours old)
      const timestamp = new Date(credentials.timestamp);
      const now = new Date();
      const hoursDiff = (now.getTime() - timestamp.getTime()) / (1000 * 60 * 60);
      
      if (hoursDiff < 24) {
        // Credentials are still valid, update UI
        updateAuthUI(true);
        showAuthStatus("Authenticated", "success");
        
        // Recreate API client with saved URL
        if (credentials.url) {
          seeqClient = new SeeqAPIClient(credentials.url);
        }
      } else {
        // Credentials expired, remove them
        localStorage.removeItem("seeq_credentials");
      }
    } catch (error) {
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
      // Silent fail for cache update
    });
  } catch (error) {
    // Silent fail for cache update
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
    
    // Check if first row contains debug information
    const firstRow = result.data[0];
    const isDebugRow = firstRow && (firstRow.Timestamp === 'DEBUG_ROW' || Object.keys(firstRow).some(key => key.startsWith('DEBUG_')));
    
    let html = `
      <h3>Data Results (${result.data.length} rows)</h3>
      <div style="background-color: #e8f4f8; padding: 10px; margin: 10px 0; border-radius: 5px;">
        <strong>Debug Info:</strong> First row Timestamp = "${firstRow?.Timestamp || 'undefined'}", 
        Has DEBUG keys: ${Object.keys(firstRow || {}).filter(key => key.startsWith('DEBUG_')).length > 0 ? 'YES' : 'NO'}
      </div>
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
    
    // Add debug information if present
    if (isDebugRow) {
      html += `
        <div style="margin-top: 20px; padding: 15px; background-color: #f0f0f0; border: 1px solid #ccc; border-radius: 5px;">
          <h4>Debug Information (First Row)</h4>
          <div style="font-family: monospace; font-size: 12px;">
      `;
      
      Object.keys(firstRow).forEach(key => {
        if (key.startsWith('DEBUG_')) {
          html += `<div><strong>${key}:</strong> ${firstRow[key]}</div>`;
        }
      });
      
      html += `
          </div>
        </div>
      `;
    }
    
    resultsDiv.innerHTML = html;
  }
}

// Display comprehensive diagnostics
function displayDiagnostics(connectionResult: any, authResult: any = null): void {
  const diagnosticsDiv = document.getElementById("diagnostics");
  if (!diagnosticsDiv) return;

  const diagnostics = seeqClient?.getDiagnostics();
  if (!diagnostics) return;

  let html = `
    <div class="diagnostics-panel">
      <h3>🔍 Connection Diagnostics</h3>
      <div class="diagnostic-section">
        <h4>Environment Information</h4>
        <div class="diagnostic-grid">
          <div class="diagnostic-item">
            <strong>App Domain:</strong> ${diagnostics.environment.appDomain}
          </div>
          <div class="diagnostic-item">
            <strong>Origin:</strong> ${diagnostics.environment.origin}
          </div>
          <div class="diagnostic-item">
            <strong>Network Type:</strong> ${diagnostics.environment.networkInfo.connectionType}
          </div>
          <div class="diagnostic-item">
            <strong>Online Status:</strong> ${diagnostics.environment.networkInfo.online ? '✅ Online' : '❌ Offline'}
          </div>
        </div>
      </div>
  `;

  // Connection test results
  if (connectionResult.diagnostics) {
    html += `
      <div class="diagnostic-section">
        <h4>Connection Test Results</h4>
        <div class="diagnostic-grid">
          <div class="diagnostic-item">
            <strong>Response Time:</strong> ${connectionResult.diagnostics.request_timing}ms
          </div>
          <div class="diagnostic-item">
            <strong>Status Code:</strong> ${connectionResult.status_code}
          </div>
          <div class="diagnostic-item">
            <strong>CORS Status:</strong> ${connectionResult.diagnostics.cors_status}
          </div>
        </div>
    `;

    // Show CORS headers if available
    if (connectionResult.diagnostics.response_headers) {
      const corsHeaders = [
        'access-control-allow-origin',
        'access-control-allow-methods',
        'access-control-allow-headers',
        'access-control-allow-credentials'
      ];
      
      html += `
        <div class="cors-headers">
          <h5>CORS Headers:</h5>
          <ul>
      `;
      
      corsHeaders.forEach(header => {
        const value = connectionResult.diagnostics.response_headers[header];
        html += `<li><strong>${header}:</strong> ${value || 'Not set'}</li>`;
      });
      
      html += `</ul></div>`;
    }
    
    html += `</div>`;
  }

  // Auth endpoint test results
  if (authResult && authResult.diagnostics) {
    html += `
      <div class="diagnostic-section">
        <h4>Authentication Endpoint Test</h4>
        <div class="diagnostic-grid">
          <div class="diagnostic-item">
            <strong>Response Time:</strong> ${authResult.diagnostics.request_timing}ms
          </div>
          <div class="diagnostic-item">
            <strong>Status Code:</strong> ${authResult.status_code}
          </div>
        </div>
    `;

    // Show CORS analysis if available
    if (authResult.diagnostics.cors_analysis) {
      const corsAnalysis = authResult.diagnostics.cors_analysis;
      html += `
        <div class="cors-analysis">
          <h5>CORS Configuration Analysis:</h5>
          <div class="cors-status">
            <div class="cors-item ${corsAnalysis.allowsOrigin ? 'success' : 'error'}">
              <strong>Origin Allowed:</strong> ${corsAnalysis.allowsOrigin ? '✅ Yes' : '❌ No'}
            </div>
            <div class="cors-item ${corsAnalysis.allowsMethods ? 'success' : 'error'}">
              <strong>Methods Allowed:</strong> ${corsAnalysis.allowsMethods ? '✅ Yes' : '❌ No'}
            </div>
            <div class="cors-item ${corsAnalysis.allowsHeaders ? 'success' : 'error'}">
              <strong>Headers Allowed:</strong> ${corsAnalysis.allowsHeaders ? '✅ Yes' : '❌ No'}
            </div>
          </div>
      `;

      if (corsAnalysis.issues.length > 0) {
        html += `
          <div class="cors-issues">
            <h6>Issues Found:</h6>
            <ul>
              ${corsAnalysis.issues.map(issue => `<li>❌ ${issue}</li>`).join('')}
            </ul>
          </div>
        `;
      }

      if (corsAnalysis.recommendations.length > 0) {
        html += `
          <div class="cors-recommendations">
            <h6>Recommendations:</h6>
            <ul>
              ${corsAnalysis.recommendations.map(rec => `<li>💡 ${rec}</li>`).join('')}
            </ul>
          </div>
        `;
      }

      html += `</div>`;
    }
    
    html += `</div>`;
  }

  // Recent diagnostic logs
  if (diagnostics.recentLogs && diagnostics.recentLogs.length > 0) {
    html += `
      <div class="diagnostic-section">
        <h4>Recent Diagnostic Logs</h4>
        <div class="log-entries">
    `;
    
    diagnostics.recentLogs.slice(-10).forEach((log: any) => {
      const timestamp = new Date(log.timestamp).toLocaleTimeString();
      html += `
        <div class="log-entry">
          <span class="log-time">${timestamp}</span>
          <span class="log-category">[${log.category}]</span>
          <span class="log-message">${log.message}</span>
        </div>
      `;
    });
    
    html += `</div></div>`;
  }

  html += `
      <div class="diagnostic-actions">
        <button onclick="clearDiagnostics()" class="btn-secondary">Clear Logs</button>
        <button onclick="exportDiagnostics()" class="btn-secondary">Export Diagnostics</button>
      </div>
    </div>
  `;

  diagnosticsDiv.innerHTML = html;
}

// Display error diagnostics
function displayErrorDiagnostics(result: any): void {
  const diagnosticsDiv = document.getElementById("diagnostics");
  if (!diagnosticsDiv) return;

  const diagnostics = seeqClient?.getDiagnostics();
  if (!diagnostics) return;

  let html = `
    <div class="diagnostics-panel error">
      <h3>❌ Error Diagnostics</h3>
      <div class="error-summary">
        <h4>Error Summary</h4>
        <div class="error-details">
          <div class="error-item">
            <strong>Error Message:</strong> ${result.message}
          </div>
          <div class="error-item">
            <strong>Error Type:</strong> ${result.error || 'Unknown'}
          </div>
        </div>
      </div>
  `;

  // Show error analysis if available
  if (result.diagnostics && result.diagnostics.error_analysis) {
    const errorAnalysis = result.diagnostics.error_analysis;
    html += `
      <div class="diagnostic-section">
        <h4>Error Analysis</h4>
        <div class="error-analysis">
          <div class="analysis-item">
            <strong>Error Type:</strong> ${errorAnalysis.errorType}
          </div>
          <div class="analysis-item">
            <strong>Likely Cause:</strong> ${errorAnalysis.likelyCause}
          </div>
          <div class="analysis-flags">
            <div class="flag ${errorAnalysis.isCorsRelated ? 'active' : ''}">CORS Related</div>
            <div class="flag ${errorAnalysis.isNetworkRelated ? 'active' : ''}">Network Related</div>
            <div class="flag ${errorAnalysis.isSslRelated ? 'active' : ''}">SSL Related</div>
            <div class="flag ${errorAnalysis.isAppDomainRelated ? 'active' : ''}">AppDomain Related</div>
          </div>
        </div>
    `;

    if (errorAnalysis.recommendations.length > 0) {
      html += `
        <div class="recommendations">
          <h5>Recommendations:</h5>
          <ul>
            ${errorAnalysis.recommendations.map(rec => `<li>💡 ${rec}</li>`).join('')}
          </ul>
        </div>
      `;
    }
    
    html += `</div>`;
  }

  // Environment information
  html += `
    <div class="diagnostic-section">
      <h4>Environment Information</h4>
      <div class="diagnostic-grid">
        <div class="diagnostic-item">
          <strong>App Domain:</strong> ${diagnostics.environment.appDomain}
        </div>
        <div class="diagnostic-item">
          <strong>Origin:</strong> ${diagnostics.environment.origin}
        </div>
        <div class="diagnostic-item">
          <strong>Network Type:</strong> ${diagnostics.environment.networkInfo.connectionType}
        </div>
        <div class="diagnostic-item">
          <strong>Online Status:</strong> ${diagnostics.environment.networkInfo.online ? '✅ Online' : '❌ Offline'}
        </div>
      </div>
    </div>
  `;

  html += `
    <div class="diagnostic-actions">
      <button onclick="clearDiagnostics()" class="btn-secondary">Clear Logs</button>
      <button onclick="exportDiagnostics()" class="btn-secondary">Export Diagnostics</button>
    </div>
  </div>
  `;

  diagnosticsDiv.innerHTML = html;
}

// Clear diagnostics
function clearDiagnostics(): void {
  if (seeqClient) {
    seeqClient.clearDiagnostics();
  }
  const diagnosticsDiv = document.getElementById("diagnostics");
  if (diagnosticsDiv) {
    diagnosticsDiv.innerHTML = '<div class="diagnostics-panel"><p>Diagnostics cleared. Run a connection test to generate new diagnostics.</p></div>';
  }
}

// Export diagnostics
function exportDiagnostics(): void {
  if (!seeqClient) return;
  
  const diagnostics = seeqClient.getDiagnostics();
  const exportData = {
    timestamp: new Date().toISOString(),
    diagnostics,
    userAgent: navigator.userAgent,
    url: window.location.href
  };
  
  const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `sqexcel-diagnostics-${new Date().toISOString().split('T')[0]}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Make functions globally available
(window as any).clearDiagnostics = clearDiagnostics;
(window as any).exportDiagnostics = exportDiagnostics;
