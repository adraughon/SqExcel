/*
 * Seeq API Client for Direct REST API Calls
 * This replaces the Python backend with direct Seeq API communication
 */

export interface SeeqAuthResult {
  success: boolean;
  message: string;
  user?: string;
  server_url?: string;
  error?: string;
  token?: string;
}

export interface SeeqConnectionResult {
  success: boolean;
  message: string;
  status_code?: number;
  error?: string;
}

export interface SeeqSensor {
  ID: string;
  Name: string;
  Type: string;
  Original_Name?: string;
  Status?: string;
}

export interface SeeqSearchResult {
  success: boolean;
  message: string;
  search_results: SeeqSensor[];
  sensor_count: number;
  error?: string;
}

export interface SeeqDataResult {
  success: boolean;
  message: string;
  search_results: SeeqSensor[];
  data: any[];
  data_columns: string[];
  data_index: string[];
  sensor_count: number;
  time_range: {
    start: string;
    end: string;
    grid: string;
  };
  error?: string;
}

export class SeeqAPIClient {
  private baseUrl: string;
  private authToken: string | null = null;
  private credentials: any = null;

  constructor(baseUrl: string) {
    this.baseUrl = baseUrl.replace(/\/$/, ''); // Remove trailing slash
  }

  /**
   * Test the authentication endpoint specifically
   */
  async testAuthEndpoint(): Promise<SeeqConnectionResult> {
    try {
      console.log(`Testing authentication endpoint: ${this.baseUrl}/api/auth/login`);
      
      // Test with OPTIONS request to check if the endpoint exists and CORS is configured
      const response = await fetch(`${this.baseUrl}/api/auth/login`, {
        method: 'OPTIONS',
        headers: {
          'Origin': window.location.origin,
          'Access-Control-Request-Method': 'POST',
          'Access-Control-Request-Headers': 'Content-Type'
        }
      });

      console.log(`Auth endpoint OPTIONS response status: ${response.status}`);
      console.log(`Auth endpoint CORS headers:`, {
        'Access-Control-Allow-Origin': response.headers.get('Access-Control-Allow-Origin'),
        'Access-Control-Allow-Methods': response.headers.get('Access-Control-Allow-Methods'),
        'Access-Control-Allow-Headers': response.headers.get('Access-Control-Allow-Headers'),
        'Access-Control-Allow-Credentials': response.headers.get('Access-Control-Allow-Credentials')
      });

      return {
        success: true,
        message: "Authentication endpoint is accessible",
        status_code: response.status
      };
    } catch (error: any) {
      console.error('Auth endpoint test error:', error);
      return {
        success: false,
        message: `Authentication endpoint test failed: ${error.message}`,
        error: error.message
      };
    }
  }

  /**
   * Test CORS preflight to the Seeq server
   */
  async testCorsPreflight(): Promise<SeeqConnectionResult> {
    try {
      console.log(`Testing CORS preflight to Seeq server: ${this.baseUrl}`);
      
      // Test with OPTIONS request to check CORS
      const response = await fetch(`${this.baseUrl}/api/system/open-ping`, {
        method: 'OPTIONS',
        headers: {
          'Origin': window.location.origin,
          'Access-Control-Request-Method': 'POST',
          'Access-Control-Request-Headers': 'Content-Type'
        }
      });

      console.log(`CORS preflight response status: ${response.status}`);
      console.log(`CORS headers:`, {
        'Access-Control-Allow-Origin': response.headers.get('Access-Control-Allow-Origin'),
        'Access-Control-Allow-Methods': response.headers.get('Access-Control-Allow-Methods'),
        'Access-Control-Allow-Headers': response.headers.get('Access-Control-Allow-Headers'),
        'Access-Control-Allow-Credentials': response.headers.get('Access-Control-Allow-Credentials')
      });

      return {
        success: true,
        message: "CORS preflight successful",
        status_code: response.status
      };
    } catch (error: any) {
      console.error('CORS preflight error:', error);
      return {
        success: false,
        message: `CORS preflight failed: ${error.message}`,
        error: error.message
      };
    }
  }

  /**
   * Test connection to Seeq server without authentication
   */
  async testConnection(): Promise<SeeqConnectionResult> {
    try {
      console.log(`Testing connection to Seeq server: ${this.baseUrl}`);
      
      // First test CORS preflight
      const corsResult = await this.testCorsPreflight();
      if (!corsResult.success) {
        console.warn('CORS preflight failed, but continuing with connection test');
      }
      
      const response = await fetch(`${this.baseUrl}/api/system/open-ping`, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
        },
      });

      console.log(`Connection test response status: ${response.status}`);
      console.log(`Response headers:`, response.headers);

      if (response.ok) {
        console.log('Connection test successful');
        return {
          success: true,
          message: "Server is reachable",
          status_code: response.status
        };
      } else {
        console.error(`Connection test failed with status: ${response.status}`);
        return {
          success: false,
          message: `Server responded with status code: ${response.status}`,
          status_code: response.status
        };
      }
    } catch (error: any) {
      console.error('Connection test error details:', error);
      
      if (error.name === 'TypeError' && error.message.includes('fetch')) {
        return {
          success: false,
          message: "Cannot connect to server - connection refused. This may be due to CORS restrictions or network issues.",
          error: "ConnectionError"
        };
      } else if (error.name === 'AbortError') {
        return {
          success: false,
          message: "Connection timeout - the server took too long to respond.",
          error: "Timeout"
        };
      } else if (error.message.includes('Failed to fetch')) {
        return {
          success: false,
          message: "Connection failed - unable to reach the server. This may be due to CORS restrictions, SSL issues, or network problems.",
          error: "FetchError"
        };
      } else if (error.message.includes('Load failed')) {
        return {
          success: false,
          message: "Load failed - the connection request could not be completed. This may be due to CORS restrictions, SSL certificate issues, or network connectivity problems.",
          error: "LoadError"
        };
      } else {
        return {
          success: false,
          message: `Connection test failed: ${error.message}`,
          error: error.message
        };
      }
    }
  }

  /**
   * Authenticate with Seeq server
   */
  async authenticate(accessKey: string, password: string, authProvider: string = 'Seeq', ignoreSslErrors: boolean = false): Promise<SeeqAuthResult> {
    try {
      // Store credentials for later use
      this.credentials = {
        accessKey,
        password,
        authProvider,
        ignoreSslErrors
      };

      console.log(`Attempting to authenticate with Seeq server: ${this.baseUrl}`);
      console.log(`Using access key: ${accessKey}, auth provider: ${authProvider}`);

      // Try to authenticate using Seeq's REST API
      const response = await fetch(`${this.baseUrl}/api/auth/login`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        credentials: 'include', // Include cookies for session management
        body: JSON.stringify({
          username: accessKey,
          password: password
        })
      });

      console.log(`Authentication response status: ${response.status}`);
      console.log(`Response headers:`, response.headers);

      if (response.ok) {
        // Seeq uses session-based authentication, so we don't get a token
        // Instead, we rely on cookies/session
        this.authToken = 'session'; // Placeholder to indicate we're authenticated
        
        console.log('Authentication successful');
        
        return {
          success: true,
          message: `Successfully authenticated as ${accessKey}`,
          user: accessKey,
          server_url: this.baseUrl,
          token: this.authToken
        };
      } else {
        const errorData = await response.json().catch(() => ({}));
        console.error('Authentication failed with response:', errorData);
        
        return {
          success: false,
          message: errorData.message || `Authentication failed with status ${response.status}`,
          error: errorData.error || `HTTP ${response.status}`
        };
      }
    } catch (error: any) {
      console.error('Authentication error details:', error);
      
      // Provide more specific error messages based on error type
      let errorMessage = 'Authentication failed';
      let errorType = 'Unknown';
      
      if (error.name === 'TypeError' && error.message.includes('fetch')) {
        errorMessage = 'Network error: Cannot connect to Seeq server. Please check the server URL and ensure it\'s accessible.';
        errorType = 'NetworkError';
      } else if (error.name === 'AbortError') {
        errorMessage = 'Request timeout: The authentication request took too long. Please try again.';
        errorType = 'TimeoutError';
      } else if (error.message.includes('Failed to fetch')) {
        errorMessage = 'Connection failed: Unable to reach the Seeq server. This may be due to CORS restrictions, network issues, or an invalid server URL.';
        errorType = 'ConnectionError';
      } else if (error.message.includes('Load failed')) {
        errorMessage = 'Load failed: The authentication request could not be completed. This may be due to CORS restrictions, SSL certificate issues, or network connectivity problems.';
        errorType = 'LoadError';
      } else {
        errorMessage = `Authentication failed: ${error.message}`;
        errorType = error.name || 'UnknownError';
      }
      
      return {
        success: false,
        message: errorMessage,
        error: errorType
      };
    }
  }

  /**
   * Search for sensors in Seeq
   */
  async searchSensors(sensorNames: string[]): Promise<SeeqSearchResult> {
    if (!this.authToken && !this.credentials) {
      return {
        success: false,
        message: "Not authenticated. Please authenticate first.",
        error: "Authentication required",
        search_results: [],
        sensor_count: 0
      };
    }

    try {
      // Use Seeq's items API to search for signals
      const searchResults: SeeqSensor[] = [];
      
      for (const sensorName of sensorNames) {
        try {
          // Search for items with the sensor name
          const response = await fetch(`${this.baseUrl}/api/items?Name=${encodeURIComponent(sensorName)}&Type=Signal`, {
            method: 'GET',
            headers: {
              'Content-Type': 'application/json',
            },
            credentials: 'include' // Include session cookies
          });

          if (response.ok) {
            const data = await response.json();
            if (data && data.length > 0) {
              // Map Seeq API response to our format
              data.forEach((result: any) => {
                searchResults.push({
                  ID: result.Id || result.id,
                  Name: result.Name || result.name,
                  Type: result.Type || 'Signal',
                  Original_Name: sensorName,
                  Status: 'Found'
                });
              });
            } else {
              // Sensor not found
              searchResults.push({
                ID: '',
                Name: sensorName,
                Type: 'Signal',
                Original_Name: sensorName,
                Status: 'Not Found'
              });
            }
          } else {
            // Search failed for this sensor
            searchResults.push({
              ID: '',
              Name: sensorName,
              Type: 'Signal',
              Original_Name: sensorName,
              Status: `Search Error: HTTP ${response.status}`
            });
          }
        } catch (error: any) {
          // Individual sensor search failed
          searchResults.push({
            ID: '',
            Name: sensorName,
            Type: 'Signal',
            Original_Name: sensorName,
            Status: `Search Error: ${error.message}`
          });
        }
      }

      const validSensors = searchResults.filter(sensor => sensor.ID);
      
      return {
        success: true,
        message: `Found ${validSensors.length} sensors`,
        search_results: searchResults,
        sensor_count: validSensors.length
      };

    } catch (error: any) {
      return {
        success: false,
        message: `Sensor search failed: ${error.message}`,
        error: error.message,
        search_results: [],
        sensor_count: 0
      };
    }
  }

  /**
   * Search for sensors and pull their data
   */
  async searchAndPullSensors(sensorNames: string[], startTime: string, endTime: string, grid: string = '15min'): Promise<SeeqDataResult> {
    try {
      // First search for sensors
      const searchResult = await this.searchSensors(sensorNames);
      
      if (!searchResult.success) {
        return {
          success: false,
          message: searchResult.message,
          error: searchResult.error,
          search_results: [],
          data: [],
          data_columns: [],
          data_index: [],
          sensor_count: 0,
          time_range: { start: startTime, end: endTime, grid }
        };
      }

      const validSensors = searchResult.search_results.filter(sensor => sensor.ID);
      
      if (validSensors.length === 0) {
        return {
          success: false,
          message: "No valid sensors found to pull data from",
          error: "All sensors failed search",
          search_results: searchResult.search_results,
          data: [],
          data_columns: [],
          data_index: [],
          sensor_count: 0,
          time_range: { start: startTime, end: endTime, grid }
        };
      }

      // Pull data for valid sensors using Seeq's signals API
      try {
        const allData: any[] = [];
        const dataColumns: string[] = [];
        
        // Get data for each sensor individually
        for (const sensor of validSensors) {
          try {
            const response = await fetch(`${this.baseUrl}/api/signals/${sensor.ID}/samples?start=${encodeURIComponent(startTime)}&end=${encodeURIComponent(endTime)}`, {
              method: 'GET',
              headers: {
                'Content-Type': 'application/json',
              },
              credentials: 'include' // Include session cookies
            });

            if (response.ok) {
              const data = await response.json();
              if (data && data.length > 0) {
                // Add sensor name as column if not already present
                const sensorColumnName = sensor.Name || `Sensor_${sensor.ID}`;
                if (!dataColumns.includes(sensorColumnName)) {
                  dataColumns.push(sensorColumnName);
                }
                
                // Add data with sensor name
                data.forEach((point: any) => {
                  const existingPoint = allData.find(p => p.Timestamp === point.Timestamp);
                  if (existingPoint) {
                    existingPoint[sensorColumnName] = point.Value;
                  } else {
                    allData.push({
                      Timestamp: point.Timestamp,
                      [sensorColumnName]: point.Value
                    });
                  }
                });
              }
            }
          } catch (error: any) {
            console.warn(`Failed to get data for sensor ${sensor.Name}:`, error);
          }
        }

        // Sort data by timestamp
        allData.sort((a, b) => new Date(a.Timestamp).getTime() - new Date(b.Timestamp).getTime());

        return {
          success: true,
          message: `Successfully retrieved data for ${validSensors.length} sensors`,
          search_results: searchResult.search_results,
          data: allData,
          data_columns: dataColumns,
          data_index: allData.map((_, i) => i.toString()),
          sensor_count: validSensors.length,
          time_range: {
            start: startTime,
            end: endTime,
            grid: grid
          }
        };
      } catch (error: any) {
        return {
          success: false,
          message: `Failed to pull data: ${error.message}`,
          error: "Data pull failed",
          search_results: searchResult.search_results,
          data: [],
          data_columns: [],
          data_index: [],
          sensor_count: validSensors.length,
          time_range: { start: startTime, end: endTime, grid }
        };
      }

    } catch (error: any) {
      return {
        success: false,
        message: `Search and pull operation failed: ${error.message}`,
        error: error.message,
        search_results: [],
        data: [],
        data_columns: [],
        data_index: [],
        sensor_count: 0,
        time_range: { start: startTime, end: endTime, grid }
      };
    }
  }

  /**
   * Get current authentication status
   */
  getAuthStatus(): { isAuthenticated: boolean; user?: string } {
    return {
      isAuthenticated: !!this.authToken,
      user: this.credentials?.accessKey
    };
  }

  /**
   * Clear authentication
   */
  logout(): void {
    this.authToken = null;
    this.credentials = null;
  }

  /**
   * Get stored credentials for persistence
   */
  getCredentials(): any {
    return this.credentials;
  }

  /**
   * Set credentials from storage
   */
  setCredentials(credentials: any): void {
    this.credentials = credentials;
  }
}
