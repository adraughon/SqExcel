const express = require('express');
const { spawn } = require('child_process');
const path = require('path');
const cors = require('cors');
const https = require('https');
const fs = require('fs');

const app = express();

// Enable CORS for Excel
app.use(cors());
app.use(express.json());

// Serve static files from the dist directory
app.use(express.static(path.join(__dirname, '../dist')));

// Backend server configuration
const PORT = 3000;

// Temporary credential storage (in production, use proper session management)
let tempCredentials = null;

// Helper function to call Python functions
function callPythonFunction(functionName, args) {
  return new Promise((resolve, reject) => {
    console.log(`[DEBUG] callPythonFunction called with function: ${functionName}`);
    console.log(`[DEBUG] Arguments:`, JSON.stringify(args, null, 2));
    
    const pythonPath = process.platform === 'win32' ? 'python' : 'python3';
    const scriptPath = path.join(__dirname, 'seeq_runner.py');
    
    console.log(`[DEBUG] Python path: ${pythonPath}`);
    console.log(`[DEBUG] Script path: ${scriptPath}`);
    
    // Create Python command with the new runner script
    const pythonArgs = [
      scriptPath,
      functionName,
      JSON.stringify(args)
    ];
    
    console.log(`[DEBUG] Python command: ${pythonPath} ${pythonArgs.join(' ')}`);
    
    const pythonProcess = spawn(pythonPath, pythonArgs);
    
    let stdout = '';
    let stderr = '';
    
    pythonProcess.stdout.on('data', (data) => {
      stdout += data.toString();
      console.log(`[DEBUG] Python stdout: ${data.toString()}`);
    });
    
    pythonProcess.stderr.on('data', (data) => {
      stderr += data.toString();
      console.log(`[DEBUG] Python stderr: ${data.toString()}`);
    });
    
    pythonProcess.on('close', (code) => {
      console.log(`[DEBUG] Python process closed with code: ${code}`);
      console.log(`[DEBUG] Final stdout: ${stdout}`);
      console.log(`[DEBUG] Final stderr: ${stderr}`);
      
      if (code === 0) {
        try {
          // Clean the output by removing ANSI escape codes and other control characters
          const cleanOutput = stdout.trim()
            .replace(/\u001b\[[0-9;]*[a-zA-Z]/g, '') // Remove ANSI escape codes
            .replace(/\r/g, '') // Remove carriage returns
            .trim();
          
          console.log(`[DEBUG] Raw Python output:`, stdout);
          console.log(`[DEBUG] Cleaned output:`, cleanOutput);
          
          const result = JSON.parse(cleanOutput);
          console.log(`[DEBUG] Successfully parsed result:`, result);
          resolve(result);
        } catch (e) {
          console.log(`[DEBUG] JSON parse error:`, e.message);
          console.log(`[DEBUG] Attempted to parse:`, stdout);
          reject(new Error(`Failed to parse Python output: ${e.message}. Raw output: ${stdout.substring(0, 200)}`));
        }
      } else {
        console.log(`[DEBUG] Python process failed with code ${code}: ${stderr}`);
        reject(new Error(`Python process failed with code ${code}: ${stderr}`));
      }
    });
    
    pythonProcess.on('error', (error) => {
      console.log(`[DEBUG] Failed to start Python process:`, error);
      reject(new Error(`Failed to start Python process: ${error.message}`));
    });
  });
}

// Authentication endpoint
app.post('/api/seeq/auth', async (req, res) => {
  try {
    const { url, accessKey, password, authProvider, ignoreSslErrors } = req.body;
    
    if (!url || !accessKey || !password) {
      return res.status(400).json({
        success: false,
        error: 'URL, access key, and password are required'
      });
    }
    
    const result = await callPythonFunction('authenticate_seeq', [
      url, accessKey, password, authProvider || 'Seeq', ignoreSslErrors === true ? 'True' : 'False'
    ]);
    
    // Store credentials for Excel functions to use (even if authentication fails)
    tempCredentials = {
      url: url,
      accessKey: accessKey,
      password: password,
      authProvider: authProvider || 'Seeq',
      ignoreSslErrors: ignoreSslErrors === true,
      timestamp: new Date().toISOString()
    };
    
    res.json(result);
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Authentication status endpoint
app.get('/api/seeq/auth/status', (req, res) => {
  if (tempCredentials) {
    res.json({
      success: true,
      isAuthenticated: true,
      message: 'Credentials available',
      credentials: tempCredentials
    });
  } else {
    res.json({
      success: true,
      isAuthenticated: false,
      message: 'Use SEEQ_AUTH function to authenticate'
    });
  }
});

// Get stored credentials endpoint
app.get('/api/seeq/credentials', (req, res) => {
  if (tempCredentials) {
    res.json({
      success: true,
      credentials: tempCredentials
    });
  } else {
    res.status(404).json({
      success: false,
      error: 'No credentials stored'
    });
  }
});

// Update stored credentials endpoint (for taskpane integration)
app.post('/api/seeq/credentials', (req, res) => {
  try {
    const { url, accessKey, password, authProvider, ignoreSslErrors, timestamp } = req.body;
    
    if (!url || !accessKey || !password) {
      return res.status(400).json({
        success: false,
        error: 'URL, access key, and password are required'
      });
    }
    
    // Store credentials for Excel functions to use
    tempCredentials = {
      url: url,
      accessKey: accessKey,
      password: password,
      authProvider: authProvider || 'Seeq',
      ignoreSslErrors: ignoreSslErrors === true,
      timestamp: timestamp || new Date().toISOString()
    };
    
    res.json({
      success: true,
      message: 'Credentials updated successfully',
      credentials: tempCredentials
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Clear stored credentials endpoint (for logout)
app.delete('/api/seeq/credentials', (req, res) => {
  try {
    tempCredentials = null;
    res.json({
      success: true,
      message: 'Credentials cleared successfully'
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Python authentication status endpoint
app.get('/api/seeq/auth/python-status', async (req, res) => {
  try {
    const result = await callPythonFunction('check_auth_status', []);
    res.json(result);
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});



// Test connection endpoint
app.post('/api/seeq/test-connection', async (req, res) => {
  try {
    const { url } = req.body;
    
    if (!url) {
      return res.status(400).json({
        success: false,
        error: 'Server URL is required'
      });
    }
    
    const result = await callPythonFunction('test_connection', [url]);
    res.json(result);
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Server info endpoint
app.post('/api/seeq/server-info', async (req, res) => {
  try {
    const { url } = req.body;
    
    if (!url) {
      return res.status(400).json({
        success: false,
        error: 'Server URL is required'
      });
    }
    
    const result = await callPythonFunction('get_server_info', [url]);
    res.json(result);
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Sensor search endpoint
app.post('/api/seeq/search-sensors', async (req, res) => {
  try {
    const { sensorNames, url, accessKey, password, authProvider, ignoreSslErrors } = req.body;
    
    if (!sensorNames || !Array.isArray(sensorNames)) {
      return res.status(400).json({
        success: false,
        error: 'Sensor names array is required'
      });
    }
    
    if (!url || !accessKey || !password) {
      return res.status(400).json({
        success: false,
        error: 'Authentication credentials are required'
      });
    }
    
    const result = await callPythonFunction('search_sensors_only', [
      sensorNames,
      url,
      accessKey,
      password,
      authProvider || 'Seeq',
      ignoreSslErrors === true ? 'True' : 'False'
    ]);
    res.json(result);
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Sensor data endpoint
app.post('/api/seeq/sensor-data', async (req, res) => {
  try {
    console.log('[DEBUG] Sensor data endpoint called');
    console.log('[DEBUG] Request body:', JSON.stringify(req.body, null, 2));
    
    const { sensorNames, startDatetime, endDatetime, grid, url, accessKey, password, authProvider, ignoreSslErrors } = req.body;
    
    if (!sensorNames || !Array.isArray(sensorNames) || !startDatetime || !endDatetime) {
      console.log('[DEBUG] Validation failed - missing required fields');
      return res.status(400).json({
        success: false,
        error: 'Sensor names, start datetime, and end datetime are required'
      });
    }
    
    if (!url || !accessKey || !password) {
      console.log('[DEBUG] Validation failed - missing authentication credentials');
      return res.status(400).json({
        success: false,
        error: 'Authentication credentials are required'
      });
    }
    
    console.log('[DEBUG] Calling Python function search_and_pull_sensors with args:', [
      sensorNames, startDatetime, endDatetime, grid || '15min', null, // timezone is null
      url,
      accessKey,
      password,
      authProvider || 'Seeq',
      ignoreSslErrors === true ? 'True' : 'False'
    ]);
    
    const result = await callPythonFunction('search_and_pull_sensors', [
      sensorNames, 
      startDatetime, 
      endDatetime, 
      grid || '15min', 
      null, // timezone is null
      url,
      accessKey,
      password,
      authProvider || 'Seeq',
      ignoreSslErrors === true ? 'True' : 'False'
    ]);
    
    console.log('[DEBUG] Python function result:', JSON.stringify(result, null, 2));
    
    res.json(result);
  } catch (error) {
    console.log('[DEBUG] Error in sensor-data endpoint:', error);
    console.log('[DEBUG] Error stack:', error.stack);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Simple test endpoint for debugging
app.get('/test', (req, res) => {
  res.json({ 
    message: 'Backend server is working!',
    timestamp: new Date().toISOString(),
    server_info: {
      node_version: process.version,
      platform: process.platform,
      arch: process.arch,
      uptime: process.uptime()
    }
  });
});

// Start server
// Try to start HTTPS server first, fallback to HTTP
try {
  // Check if we have SSL certificates (for development)
  const httpsOptions = {
    key: fs.readFileSync(path.join(process.env.HOME || process.env.USERPROFILE, '.office-addin-dev-certs/localhost.key')),
    cert: fs.readFileSync(path.join(process.env.HOME || process.env.USERPROFILE, '.office-addin-dev-certs/localhost.crt'))
  };
  
  https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`TSFlow backend server running on HTTPS port ${PORT}`);
    console.log(`Health check: https://localhost:${PORT}/health`);
  });
} catch (error) {
  console.log('HTTPS certificates not found, starting HTTP server...');
  app.listen(PORT, () => {
    console.log(`TSFlow backend server running on HTTP port ${PORT}`);
    console.log(`Health check: http://localhost:${PORT}/health`);
  });
}

module.exports = app;
