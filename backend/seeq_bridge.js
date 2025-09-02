const { spawn } = require('child_process');
const path = require('path');

class SeeqBridge {
    constructor() {
        this.pythonPath = 'python3'; // Default to python3
        this.scriptPath = path.join(__dirname, 'seeq_auth.py');
        this.isAuthenticated = false;
        this.authInfo = null;
    }

    /**
     * Test connection to Seeq server
     * @param {string} url - Seeq server URL
     * @returns {Promise<Object>} Connection test result
     */
    async testConnection(url) {
        return new Promise((resolve, reject) => {
            const pythonProcess = spawn(this.pythonPath, [
                '-c',
                `import sys; sys.path.append('${__dirname}'); from seeq_auth import test_connection; import json; result = test_connection('${url}'); print(json.dumps(result))`
            ]);

            let output = '';
            let errorOutput = '';

            pythonProcess.stdout.on('data', (data) => {
                output += data.toString();
            });

            pythonProcess.stderr.on('data', (data) => {
                errorOutput += data.toString();
            });

            pythonProcess.on('close', (code) => {
                if (code === 0) {
                    try {
                        const result = JSON.parse(output.trim());
                        resolve(result);
                    } catch (e) {
                        reject(new Error(`Failed to parse Python output: ${e.message}`));
                    }
                } else {
                    reject(new Error(`Python process failed with code ${code}: ${errorOutput}`));
                }
            });

            pythonProcess.on('error', (error) => {
                reject(new Error(`Failed to start Python process: ${error.message}`));
            });
        });
    }

    /**
     * Authenticate with Seeq server
     * @param {string} url - Seeq server URL
     * @param {string} accessKey - Seeq access key
     * @param {string} password - Seeq password
     * @param {string} authProvider - Authentication provider (default: 'Seeq')
     * @param {boolean} ignoreSslErrors - Whether to ignore SSL errors
     * @returns {Promise<Object>} Authentication result
     */
    async authenticate(url, accessKey, password, authProvider = 'Seeq', ignoreSslErrors = false) {
        return new Promise((resolve, reject) => {
            const pythonProcess = spawn(this.pythonPath, [
                '-c',
                `import sys; sys.path.append('${__dirname}'); from seeq_auth import authenticate_seeq; import json; result = authenticate_seeq('${url}', '${accessKey}', '${password}', '${authProvider}', ${ignoreSslErrors}); print(json.dumps(result))`
            ]);

            let output = '';
            let errorOutput = '';

            pythonProcess.stdout.on('data', (data) => {
                output += data.toString();
            });

            pythonProcess.stderr.on('data', (data) => {
                errorOutput += data.toString();
            });

            pythonProcess.on('close', (code) => {
                if (code === 0) {
                    try {
                        const result = JSON.parse(output.trim());
                        if (result.success) {
                            this.isAuthenticated = true;
                            this.authInfo = result;
                        }
                        resolve(result);
                    } catch (e) {
                        reject(new Error(`Failed to parse Python output: ${e.message}`));
                    }
                } else {
                    reject(new Error(`Python process failed with code ${code}: ${errorOutput}`));
                }
            });

            pythonProcess.on('error', (error) => {
                reject(new Error(`Failed to start Python process: ${error.message}`));
            });
        });
    }

    /**
     * Get server information
     * @param {string} url - Seeq server URL
     * @returns {Promise<Object>} Server information
     */
    async getServerInfo(url) {
        return new Promise((resolve, reject) => {
            const pythonProcess = spawn(this.pythonPath, [
                '-c',
                `import sys; sys.path.append('${__dirname}'); from seeq_auth import get_server_info; import json; result = get_server_info('${url}'); print(json.dumps(result))`
            ]);

            let output = '';
            let errorOutput = '';

            pythonProcess.stdout.on('data', (data) => {
                output += data.toString();
            });

            pythonProcess.stderr.on('data', (data) => {
                errorOutput += data.toString();
            });

            pythonProcess.on('close', (code) => {
                if (code === 0) {
                    try {
                        const result = JSON.parse(output.trim());
                        resolve(result);
                    } catch (e) {
                        reject(new Error(`Failed to parse Python output: ${e.message}`));
                    }
                } else {
                    reject(new Error(`Python process failed with code ${code}: ${errorOutput}`));
                }
            });

            pythonProcess.on('error', (error) => {
                reject(new Error(`Failed to start Python process: ${error.message}`));
            });
        });
    }

    /**
     * Search for sensors and pull their data
     * @param {Array<string>} sensorNames - Array of sensor names to search for
     * @param {string} startDatetime - Start time for data pull (ISO format)
     * @param {string} endDatetime - End time for data pull (ISO format)
     * @param {string} grid - Grid interval for data (e.g., '15min', '1h')
     * @param {string} timezone - Optional timezone for datetime parsing
     * @returns {Promise<Object>} Search and pull results
     */
    async searchAndPullSensors(sensorNames, startDatetime, endDatetime, grid = '15min', timezone = null) {
        return new Promise((resolve, reject) => {
            const sensorNamesJson = JSON.stringify(sensorNames);
            const timezoneParam = timezone ? `, '${timezone}'` : '';
            
            const pythonProcess = spawn(this.pythonPath, [
                '-c',
                `import sys; sys.path.append('${__dirname}'); from seeq_auth import search_and_pull_sensors; import json; result = search_and_pull_sensors(${sensorNamesJson}, '${startDatetime}', '${endDatetime}', '${grid}'${timezoneParam}); print(json.dumps(result))`
            ]);

            let output = '';
            let errorOutput = '';

            pythonProcess.stdout.on('data', (data) => {
                output += data.toString();
            });

            pythonProcess.stderr.on('data', (data) => {
                errorOutput += data.toString();
            });

            pythonProcess.on('close', (code) => {
                if (code === 0) {
                    try {
                        const result = JSON.parse(output.trim());
                        resolve(result);
                    } catch (e) {
                        reject(new Error(`Failed to parse Python output: ${e.message}`));
                    }
                } else {
                    reject(new Error(`Python process failed with code ${code}: ${errorOutput}`));
                }
            });

            pythonProcess.on('error', (error) => {
                reject(new Error(`Failed to start Python process: ${error.message}`));
            });
        });
    }

    /**
     * Search for sensors only (without pulling data)
     * @param {Array<string>} sensorNames - Array of sensor names to search for
     * @returns {Promise<Object>} Search results
     */
    async searchSensorsOnly(sensorNames) {
        return new Promise((resolve, reject) => {
            const sensorNamesJson = JSON.stringify(sensorNames);
            
            const pythonProcess = spawn(this.pythonPath, [
                '-c',
                `import sys; sys.path.append('${__dirname}'); from seeq_auth import search_sensors_only; import json; result = search_sensors_only(${sensorNamesJson}); print(json.dumps(result))`
            ]);

            let output = '';
            let errorOutput = '';

            pythonProcess.stdout.on('data', (data) => {
                output += data.toString();
            });

            pythonProcess.stderr.on('data', (data) => {
                errorOutput += data.toString();
            });

            pythonProcess.on('close', (code) => {
                if (code === 0) {
                    try {
                        const result = JSON.parse(output.trim());
                        resolve(result);
                    } catch (e) {
                        reject(new Error(`Failed to parse Python output: ${e.message}`));
                    }
                } else {
                    reject(new Error(`Python process failed with code ${code}: ${errorOutput}`));
                }
            });

            pythonProcess.on('error', (error) => {
                reject(new Error(`Failed to start Python process: ${error.message}`));
            });
        });
    }

    /**
     * Check if currently authenticated
     * @returns {boolean} Authentication status
     */
    isCurrentlyAuthenticated() {
        return this.isAuthenticated;
    }

    /**
     * Get current authentication info
     * @returns {Object|null} Current authentication information
     */
    getCurrentAuthInfo() {
        return this.authInfo;
    }

    /**
     * Logout (clear authentication state)
     */
    logout() {
        this.isAuthenticated = false;
        this.authInfo = null;
    }
}

module.exports = SeeqBridge;
