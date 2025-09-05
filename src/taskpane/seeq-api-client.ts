/*
 * Seeq API Client for Proxy Server (SqExcelWeb)
 * This client communicates with the FastAPI proxy server to avoid CORS issues
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
  diagnostics?: {
    cors_headers?: any;
    response_headers?: any;
    request_timing?: number;
    user_agent?: string;
    origin?: string;
    app_domain?: string;
    network_type?: string;
    ssl_info?: any;
    cors_analysis?: any;
    error_analysis?: any;
    cors_status?: string;
  };
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
  private proxyUrl: string;
  private seeqServerUrl: string;
  private authToken: string | null = null;
  private credentials: any = null;
  private diagnosticLog: any[] = [];

  constructor(seeqServerUrl: string) {
    this.seeqServerUrl = seeqServerUrl.replace(/\/$/, ''); // Remove trailing slash
    this.proxyUrl = 'https://sq-excel-web.vercel.app';
    this.logDiagnostic('CLIENT_INIT', `SeeqAPIClient initialized with proxy: ${this.proxyUrl}, Seeq server: ${this.seeqServerUrl}`);
  }

  /**
   * Log diagnostic information for debugging
   */
  private logDiagnostic(category: string, message: string, data?: any): void {
    const logEntry = {
      timestamp: new Date().toISOString(),
      category,
      message,
      data,
      userAgent: navigator.userAgent,
      origin: window.location.origin,
      appDomain: this.detectAppDomain(),
      networkInfo: this.getNetworkInfo()
    };
    
    this.diagnosticLog.push(logEntry);
    console.log(`[${category}] ${message}`, data || '');
  }

  /**
   * Detect the current AppDomain context
   */
  private detectAppDomain(): string {
    try {
      if (typeof Office !== 'undefined' && Office.context) {
        return `Office.js - ${Office.context.host || 'Unknown Host'}`;
      }
      if (window.location.protocol === 'https:') {
        return 'HTTPS Web Context';
      }
      if (window.location.protocol === 'http:') {
        return 'HTTP Web Context';
      }
      if (window.location.protocol === 'file:') {
        return 'File Protocol Context';
      }
      return `Unknown Context - ${window.location.protocol}`;
    } catch (error) {
      return `Error detecting context: ${error}`;
    }
  }

  /**
   * Get network information
   */
  private getNetworkInfo(): any {
    try {
      const connection = (navigator as any).connection || (navigator as any).mozConnection || (navigator as any).webkitConnection;
      return {
        online: navigator.onLine,
        connectionType: connection?.effectiveType || 'unknown',
        downlink: connection?.downlink || 'unknown',
        rtt: connection?.rtt || 'unknown'
      };
    } catch (error) {
      return { online: navigator.onLine, error: error.toString() };
    }
  }

  /**
   * Get comprehensive diagnostic information
   */
  getDiagnostics(): any {
    return {
      clientInfo: {
        proxyUrl: this.proxyUrl,
        seeqServerUrl: this.seeqServerUrl,
        isAuthenticated: !!this.authToken,
        diagnosticLogCount: this.diagnosticLog.length
      },
      environment: {
        userAgent: navigator.userAgent,
        origin: window.location.origin,
        appDomain: this.detectAppDomain(),
        networkInfo: this.getNetworkInfo(),
        timestamp: new Date().toISOString()
      },
      recentLogs: this.diagnosticLog.slice(-20) // Last 20 log entries
    };
  }

  /**
   * Clear diagnostic logs
   */
  clearDiagnostics(): void {
    this.diagnosticLog = [];
    this.logDiagnostic('DIAGNOSTICS_CLEARED', 'Diagnostic logs cleared');
  }

  /**
   * Test connection to the proxy server
   */
  async testConnection(): Promise<SeeqConnectionResult> {
    const startTime = Date.now();
    const url = `${this.proxyUrl}/api/seeq/test-connection`;
    
    this.logDiagnostic('PROXY_CONNECTION_TEST_START', `Testing connection to proxy server: ${this.proxyUrl}`, {
      url,
      seeqServerUrl: this.seeqServerUrl,
      appDomain: this.detectAppDomain(),
      networkInfo: this.getNetworkInfo()
    });
    
    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          seeq_url: this.seeqServerUrl
        })
      });

      const requestTime = Date.now() - startTime;
      
      // Log all response headers
      const allHeaders: any = {};
      response.headers.forEach((value, key) => {
        allHeaders[key] = value;
      });

      this.logDiagnostic('PROXY_CONNECTION_RESPONSE', `Proxy connection test response received`, {
        status: response.status,
        statusText: response.statusText,
        requestTime,
        allHeaders,
        url: response.url
      });

      if (response.ok) {
        const data = await response.json();
        this.logDiagnostic('PROXY_CONNECTION_SUCCESS', 'Proxy connection test successful', data);
        
        return {
          success: true,
          message: data.message || "Proxy server is reachable and can connect to Seeq",
          status_code: response.status,
          diagnostics: {
            response_headers: allHeaders,
            request_timing: requestTime,
            user_agent: navigator.userAgent,
            origin: window.location.origin,
            app_domain: this.detectAppDomain(),
            network_type: this.getNetworkInfo().connectionType,
            cors_status: 'OK'
          }
        };
      } else {
        const errorData = await response.json().catch(() => ({}));
        this.logDiagnostic('PROXY_CONNECTION_FAILED', `Proxy connection test failed with status: ${response.status}`, {
          status: response.status,
          statusText: response.statusText,
          headers: allHeaders,
          errorData
        });
        
        return {
          success: false,
          message: errorData.message || `Proxy server responded with status code: ${response.status}`,
          status_code: response.status,
          error: errorData.error || `HTTP ${response.status}`,
          diagnostics: {
            response_headers: allHeaders,
            request_timing: requestTime,
            user_agent: navigator.userAgent,
            origin: window.location.origin,
            app_domain: this.detectAppDomain(),
            network_type: this.getNetworkInfo().connectionType,
            cors_status: 'OK'
          }
        };
      }
    } catch (error: any) {
      const requestTime = Date.now() - startTime;
      
      this.logDiagnostic('PROXY_CONNECTION_ERROR', `Proxy connection test error`, {
        error: error.message,
        errorName: error.name,
        errorStack: error.stack,
        requestTime,
        url
      });
      
      let message = "Proxy connection test failed";
      let errorType = "Unknown";
      
      if (error.name === 'TypeError' && error.message.includes('fetch')) {
        message = "Cannot connect to proxy server - connection refused. This may be due to network issues.";
        errorType = "ConnectionError";
      } else if (error.name === 'AbortError') {
        message = "Connection timeout - the proxy server took too long to respond.";
        errorType = "Timeout";
      } else if (error.message.includes('Failed to fetch')) {
        message = "Connection failed - unable to reach the proxy server. This may be due to network problems.";
        errorType = "FetchError";
      } else {
        message = `Proxy connection test failed: ${error.message}`;
        errorType = error.message;
      }
      
      return {
        success: false,
        message,
        error: errorType,
        diagnostics: {
          request_timing: requestTime,
          user_agent: navigator.userAgent,
          origin: window.location.origin,
          app_domain: this.detectAppDomain(),
          network_type: this.getNetworkInfo().connectionType,
          cors_status: 'OK'
        }
      };
    }
  }

  /**
   * Authenticate with Seeq server through proxy
   */
  async authenticate(accessKey: string, password: string, authProvider: string = 'Seeq', ignoreSslErrors: boolean = false): Promise<SeeqAuthResult> {
    try {
      // Store credentials for later use
      this.credentials = {
        accessKey,
        password,
        authProvider,
        ignoreSslErrors,
        seeq_url: this.seeqServerUrl
      };

      this.logDiagnostic('PROXY_AUTH_START', `Attempting to authenticate through proxy: ${this.proxyUrl}`, {
        seeqServerUrl: this.seeqServerUrl,
        accessKey,
        authProvider
      });

      const response = await fetch(`${this.proxyUrl}/api/seeq/auth`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          seeq_url: this.seeqServerUrl,
          username: accessKey,
          password: password,
          auth_provider: authProvider,
          ignore_ssl_errors: ignoreSslErrors
        })
      });

      this.logDiagnostic('PROXY_AUTH_RESPONSE', `Proxy authentication response received`, {
        status: response.status,
        statusText: response.statusText
      });

      if (response.ok) {
        const data = await response.json();
        this.authToken = data.token || 'authenticated';
        
        this.logDiagnostic('PROXY_AUTH_SUCCESS', 'Proxy authentication successful', {
          user: data.user,
          server_url: data.server_url
        });
        
        return {
          success: true,
          message: data.message || `Successfully authenticated as ${accessKey}`,
          user: data.user || accessKey,
          server_url: data.server_url || this.seeqServerUrl,
          token: this.authToken
        };
      } else {
        const errorData = await response.json().catch(() => ({}));
        this.logDiagnostic('PROXY_AUTH_FAILED', 'Proxy authentication failed', {
          status: response.status,
          errorData
        });
        
        return {
          success: false,
          message: errorData.message || `Authentication failed with status ${response.status}`,
          error: errorData.error || `HTTP ${response.status}`
        };
      }
    } catch (error: any) {
      this.logDiagnostic('PROXY_AUTH_ERROR', 'Proxy authentication error', {
        error: error.message,
        errorName: error.name
      });
      
      let errorMessage = 'Authentication failed';
      let errorType = 'Unknown';
      
      if (error.name === 'TypeError' && error.message.includes('fetch')) {
        errorMessage = 'Network error: Cannot connect to proxy server. Please check your internet connection.';
        errorType = 'NetworkError';
      } else if (error.name === 'AbortError') {
        errorMessage = 'Request timeout: The authentication request took too long. Please try again.';
        errorType = 'TimeoutError';
      } else if (error.message.includes('Failed to fetch')) {
        errorMessage = 'Connection failed: Unable to reach the proxy server. This may be due to network issues.';
        errorType = 'ConnectionError';
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
   * Search for sensors in Seeq through proxy
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
      this.logDiagnostic('PROXY_SEARCH_START', `Searching for sensors through proxy`, {
        sensorNames,
        seeqServerUrl: this.seeqServerUrl
      });

      const response = await fetch(`${this.proxyUrl}/api/seeq/search`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          seeq_url: this.seeqServerUrl,
          sensor_names: sensorNames,
          username: this.credentials?.accessKey,
          password: this.credentials?.password,
          auth_provider: this.credentials?.authProvider || 'Seeq'
        })
      });

      this.logDiagnostic('PROXY_SEARCH_RESPONSE', `Proxy search response received`, {
        status: response.status,
        statusText: response.statusText
      });

      if (response.ok) {
        const data = await response.json();
        this.logDiagnostic('PROXY_SEARCH_SUCCESS', 'Proxy search successful', {
          sensorCount: data.sensor_count,
          searchResults: data.search_results
        });

        return {
          success: true,
          message: data.message || `Found ${data.sensor_count} sensors`,
          search_results: data.search_results || [],
          sensor_count: data.sensor_count || 0
        };
      } else {
        const errorData = await response.json().catch(() => ({}));
        this.logDiagnostic('PROXY_SEARCH_FAILED', 'Proxy search failed', {
          status: response.status,
          errorData
        });

        return {
          success: false,
          message: errorData.message || `Search failed with status ${response.status}`,
          error: errorData.error || `HTTP ${response.status}`,
          search_results: [],
          sensor_count: 0
        };
      }
    } catch (error: any) {
      this.logDiagnostic('PROXY_SEARCH_ERROR', 'Proxy search error', {
        error: error.message,
        errorName: error.name
      });

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
   * Search for sensors and pull their data through proxy
   */
  async searchAndPullSensors(sensorNames: string[], startTime: string, endTime: string, grid: string = '15min'): Promise<SeeqDataResult> {
    try {
      this.logDiagnostic('PROXY_SEARCH_PULL_START', `Searching and pulling sensor data through proxy`, {
        sensorNames,
        startTime,
        endTime,
        grid,
        seeqServerUrl: this.seeqServerUrl
      });

      const response = await fetch(`${this.proxyUrl}/api/seeq/data`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          seeq_url: this.seeqServerUrl,
          sensor_names: sensorNames,
          start_time: startTime,
          end_time: endTime,
          grid: grid,
          username: this.credentials?.accessKey,
          password: this.credentials?.password,
          auth_provider: this.credentials?.authProvider || 'Seeq'
        })
      });

      this.logDiagnostic('PROXY_SEARCH_PULL_RESPONSE', `Proxy search and pull response received`, {
        status: response.status,
        statusText: response.statusText
      });

      if (response.ok) {
        const data = await response.json();
        this.logDiagnostic('PROXY_SEARCH_PULL_SUCCESS', 'Proxy search and pull successful', {
          sensorCount: data.sensor_count,
          dataLength: data.data?.length || 0,
          dataColumns: data.data_columns
        });

        return {
          success: true,
          message: data.message || `Successfully retrieved data for ${data.sensor_count} sensors`,
          search_results: data.search_results || [],
          data: data.data || [],
          data_columns: data.data_columns || [],
          data_index: data.data_index || [],
          sensor_count: data.sensor_count || 0,
          time_range: {
            start: startTime,
            end: endTime,
            grid: grid
          }
        };
      } else {
        const errorData = await response.json().catch(() => ({}));
        this.logDiagnostic('PROXY_SEARCH_PULL_FAILED', 'Proxy search and pull failed', {
          status: response.status,
          errorData
        });

        return {
          success: false,
          message: errorData.message || `Search and pull failed with status ${response.status}`,
          error: errorData.error || `HTTP ${response.status}`,
          search_results: [],
          data: [],
          data_columns: [],
          data_index: [],
          sensor_count: 0,
          time_range: { start: startTime, end: endTime, grid }
        };
      }
    } catch (error: any) {
      this.logDiagnostic('PROXY_SEARCH_PULL_ERROR', 'Proxy search and pull error', {
        error: error.message,
        errorName: error.name
      });

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
  async logout(): Promise<void> {
    try {
      if (this.authToken) {
        this.logDiagnostic('PROXY_LOGOUT_START', 'Logging out through proxy');
        
        const response = await fetch(`${this.proxyUrl}/api/seeq/auth`, {
          method: 'DELETE',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            seeq_url: this.seeqServerUrl,
            username: this.credentials?.accessKey
          })
        });

        this.logDiagnostic('PROXY_LOGOUT_RESPONSE', `Proxy logout response received`, {
          status: response.status,
          statusText: response.statusText
        });
      }
    } catch (error: any) {
      this.logDiagnostic('PROXY_LOGOUT_ERROR', 'Proxy logout error', {
        error: error.message
      });
    } finally {
      this.authToken = null;
      this.credentials = null;
    }
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
