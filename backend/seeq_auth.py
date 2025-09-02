#!/usr/bin/env python3
"""
Seeq Authentication Backend using SPy
This module provides authentication functionality for the TSFlow Excel extension
"""

import json
import sys
import traceback
from typing import Dict, Any, Optional

# SPy is imported in seeq_runner.py - this file should not import it separately
# to avoid conflicts. All SPy operations should go through seeq_runner.py

def authenticate_seeq(url: str, access_key: str, password: str, 
                     auth_provider: str = 'Seeq', 
                     ignore_ssl_errors: any = False) -> Dict[str, Any]:
    """
    Authenticate with Seeq server using SPy
    
    Args:
        url: Seeq server URL
        access_key: Seeq access key
        password: Seeq password
        auth_provider: Authentication provider (default: 'Seeq')
        ignore_ssl_errors: Whether to ignore SSL certificate errors
        
    Returns:
        Dictionary containing authentication result and status
    """
    try:
        # Suppress SPy output by redirecting stdout temporarily
        import io
        import sys
        from contextlib import redirect_stdout
        
        # Set compatibility option for maximum compatibility
        # Use a version that exists in the available package
        try:
            spy.options.compatibility = 66
        except AttributeError:
            # If compatibility option doesn't exist, continue without it
            pass
        
        # Convert ignore_ssl_errors to proper boolean
        if isinstance(ignore_ssl_errors, str):
            ignore_ssl_errors = ignore_ssl_errors.lower() in ('true', '1', 'yes', 'on')
        elif not isinstance(ignore_ssl_errors, bool):
            ignore_ssl_errors = False
        
        # Suppress SPy output during login
        with redirect_stdout(io.StringIO()):
            # Attempt to login
            spy.login(
                url=url,
                access_key=access_key,
                password=password,
                ignore_ssl_errors=ignore_ssl_errors
            )
        
        # Check if login was successful
        if spy.user is not None:
            return {
                "success": True,
                "message": f"Successfully authenticated as {spy.user}",
                "user": str(spy.user),
                "server_url": url
            }
        else:
            return {
                "success": False,
                "message": "Authentication failed - no user returned",
                "error": "No user returned from SPy"
            }
            
    except Exception as e:
        error_msg = str(e)
        error_trace = traceback.format_exc()
        
        return {
            "success": False,
            "message": f"Authentication failed: {error_msg}",
            "error": error_msg,
            "traceback": error_trace
        }

def test_connection(url: str) -> Dict[str, Any]:
    """
    Test connection to Seeq server without authentication
    
    Args:
        url: Seeq server URL to test
        
    Returns:
        Dictionary containing connection test result
    """
    try:
        import requests
        
        # Test basic connectivity
        response = requests.get(f"{url}/api/status", timeout=10)
        
        if response.status_code == 200:
            return {
                "success": True,
                "message": "Server is reachable",
                "status_code": response.status_code
            }
        else:
            return {
                "success": False,
                "message": f"Server responded with status code: {response.status_code}",
                "status_code": response.status_code
            }
            
    except requests.exceptions.ConnectionError:
        return {
            "success": False,
            "message": "Cannot connect to server - connection refused",
            "error": "ConnectionError"
        }
    except requests.exceptions.Timeout:
        return {
            "success": False,
            "message": "Connection timeout",
            "error": "Timeout"
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"Connection test failed: {str(e)}",
            "error": str(e)
        }

def get_server_info(url: str) -> Dict[str, Any]:
    """
    Get basic information about the Seeq server
    
    Args:
        url: Seeq server URL
        
    Returns:
        Dictionary containing server information
    """
    try:
        import requests
        
        response = requests.get(f"{url}/api/status", timeout=10)
        
        if response.status_code == 200:
            try:
                data = response.json()
                return {
                    "success": True,
                    "server_info": data,
                    "message": "Server information retrieved successfully"
                }
            except json.JSONDecodeError:
                return {
                    "success": True,
                    "server_info": {"raw_response": response.text},
                    "message": "Server responded but response is not JSON"
                }
        else:
            return {
                "success": False,
                "message": f"Failed to get server info: {response.status_code}",
                "status_code": response.status_code
            }
            
    except Exception as e:
        return {
            "success": False,
            "message": f"Failed to get server info: {str(e)}",
            "error": str(e)
        }

def check_auth_status() -> Dict[str, Any]:
    """
    Check the current authentication status with Seeq
    
    Returns:
        Dictionary containing authentication status
    """
    try:
        if spy.user is not None:
            return {
                "success": True,
                "isAuthenticated": True,
                "user": str(spy.user),
                "message": f"Authenticated as {spy.user}"
            }
        else:
            return {
                "success": True,
                "isAuthenticated": False,
                "message": "Not authenticated"
            }
    except Exception as e:
        return {
            "success": False,
            "isAuthenticated": False,
            "error": str(e),
            "message": "Error checking authentication status"
        }

def search_and_pull_sensors(sensor_names: list, start_datetime: str, end_datetime: str, 
                           grid: str = '15min', timezone: str = None, url: str = None, 
                           access_key: str = None, password: str = None, 
                           auth_provider: str = 'Seeq', ignore_ssl_errors: any = False) -> Dict[str, Any]:
    """
    Search for sensors in Seeq and pull their data
    
    Args:
        sensor_names: List of sensor names to search for
        start_datetime: Start time for data pull (ISO format string)
        end_datetime: End time for data pull (ISO format string)
        grid: Grid interval for data (e.g., '15min', '1h', '1d')
        timezone: Timezone for datetime parsing (defaults to system timezone)
        url: Seeq server URL (for re-authentication if needed)
        access_key: Seeq access key (for re-authentication if needed)
        password: Seeq password (for re-authentication if needed)
        auth_provider: Authentication provider (for re-authentication if needed)
        ignore_ssl_errors: Whether to ignore SSL errors (for re-authentication if needed)
        
    Returns:
        Dictionary containing search results and data
    """
    try:
        import pandas as pd
        from datetime import datetime
        
        # Check if we're authenticated, if not, try to authenticate
        if spy.user is None:
            if url and access_key and password:
                # Try to authenticate first
                auth_result = authenticate_seeq(url, access_key, password, auth_provider, ignore_ssl_errors)
                if not auth_result['success']:
                    return {
                        "success": False,
                        "message": "Authentication required and re-authentication failed",
                        "error": "Authentication required"
                    }
            else:
                return {
                    "success": False,
                    "message": "Not authenticated to Seeq. Please login first.",
                    "error": "Authentication required"
                }
        
        # Parse datetime strings
        try:
            if timezone:
                start_dt = pd.to_datetime(start_datetime, utc=True).tz_convert(timezone)
                end_dt = pd.to_datetime(end_datetime, utc=True).tz_convert(timezone)
            else:
                start_dt = pd.to_datetime(start_datetime)
                end_dt = pd.to_datetime(end_datetime)
        except Exception as e:
            return {
                "success": False,
                "message": f"Invalid datetime format: {str(e)}",
                "error": "Datetime parsing failed"
            }
        
        # Search for sensors
        search_results = []
        for sensor_name in sensor_names:
            try:
                # Search with Type set to StoredSignal
                result = spy.search({
                    'Name': sensor_name,
                    'Type': 'StoredSignal'
                })
                
                if not result.empty:
                    # Add the sensor name for reference
                    result['Original_Name'] = sensor_name
                    search_results.append(result)
                else:
                    # Create a placeholder for sensors not found
                    placeholder = pd.DataFrame([{
                        'Name': sensor_name,
                        'ID': None,
                        'Type': 'StoredSignal',
                        'Original_Name': sensor_name,
                        'Status': 'Not Found'
                    }])
                    search_results.append(placeholder)
                    
            except Exception as e:
                # Create error placeholder
                error_placeholder = pd.DataFrame([{
                    'Name': sensor_name,
                    'ID': None,
                    'Type': 'StoredSignal',
                    'Original_Name': sensor_name,
                    'Status': f'Search Error: {str(e)}'
                }])
                search_results.append(error_placeholder)
        
        # Combine all search results
        if search_results:
            combined_results = pd.concat(search_results, ignore_index=True)
        else:
            return {
                "success": False,
                "message": "No sensors found or search failed",
                "error": "Search returned no results"
            }
        
        # Filter to only sensors that were found successfully
        valid_sensors = combined_results[combined_results['ID'].notna()]
        
        if valid_sensors.empty:
            return {
                "success": False,
                "message": "No valid sensors found to pull data from",
                "error": "All sensors failed search",
                "search_results": combined_results.to_dict('records')
            }
        
        # Pull data for valid sensors
        try:
            data_df = spy.pull(
                valid_sensors,
                start=start_dt,
                end=end_dt,
                grid=grid,
                header='Name'
            )
            
            # Convert to records for JSON serialization
            data_records = data_df.reset_index().to_dict('records')
            
            return {
                "success": True,
                "message": f"Successfully retrieved data for {len(valid_sensors)} sensors",
                "search_results": combined_results.to_dict('records'),
                "data": data_records,
                "data_columns": list(data_df.columns),
                "data_index": [str(idx) for idx in data_df.index],
                "sensor_count": len(valid_sensors),
                "time_range": {
                    "start": str(start_dt),
                    "end": str(end_dt),
                    "grid": grid
                }
            }
            
        except Exception as e:
            return {
                "success": False,
                "message": f"Failed to pull data: {str(e)}",
                "error": "Data pull failed",
                "search_results": combined_results.to_dict('records')
            }
            
    except Exception as e:
        error_msg = str(e)
        error_trace = traceback.format_exc()
        
        return {
            "success": False,
            "message": f"Search and pull operation failed: {error_msg}",
            "error": error_msg,
            "traceback": error_trace
        }

def search_sensors_only(sensor_names: list, url: str = None, access_key: str = None, 
                       password: str = None, auth_provider: str = 'Seeq', 
                       ignore_ssl_errors: any = False) -> Dict[str, Any]:
    """
    Search for sensors in Seeq without pulling data
    
    Args:
        sensor_names: List of sensor names to search for
        url: Seeq server URL (for re-authentication if needed)
        access_key: Seeq access key (for re-authentication if needed)
        password: Seeq password (for re-authentication if needed)
        auth_provider: Authentication provider (for re-authentication if needed)
        ignore_ssl_errors: Whether to ignore SSL errors (for re-authentication if needed)
        
    Returns:
        Dictionary containing search results
    """
    try:
        # Check if we're authenticated, if not, try to authenticate
        if spy.user is None:
            if url and access_key and password:
                # Try to authenticate first
                auth_result = authenticate_seeq(url, access_key, password, auth_provider, ignore_ssl_errors)
                if not auth_result['success']:
                    return {
                        "success": False,
                        "message": "Authentication required and re-authentication failed",
                        "error": "Authentication required"
                    }
            else:
                return {
                    "success": False,
                    "message": "Not authenticated to Seeq. Please login first.",
                    "error": "Authentication required"
                }
        
        # Search for sensors
        search_results = []
        for sensor_name in sensor_names:
            try:
                # Search with Type set to StoredSignal
                result = spy.search({
                    'Name': sensor_name,
                    'Type': 'StoredSignal'
                })
                
                if not result.empty:
                    # Add the sensor name for reference
                    result['Original_Name'] = sensor_name
                    search_results.append(result)
                else:
                    # Create a placeholder for sensors not found
                    placeholder = pd.DataFrame([{
                        'Name': sensor_name,
                        'ID': None,
                        'Type': 'StoredSignal',
                        'Original_Name': sensor_name,
                        'Status': 'Not Found'
                    }])
                    search_results.append(placeholder)
                    
            except Exception as e:
                # Create error placeholder
                error_placeholder = pd.DataFrame([{
                    'Name': sensor_name,
                    'ID': None,
                    'Type': 'StoredSignal',
                    'Original_Name': sensor_name,
                    'Status': f'Search Error: {str(e)}'
                }])
                search_results.append(error_placeholder)
        
        # Combine all search results
        if search_results:
            combined_results = pd.concat(search_results, ignore_index=True)
            return {
                "success": True,
                "message": f"Search completed for {len(sensor_names)} sensors",
                "search_results": combined_results.to_dict('records'),
                "sensor_count": len(sensor_names)
            }
        else:
            return {
                "success": False,
                "message": "Search failed for all sensors",
                "error": "No search results"
            }
            
    except Exception as e:
        error_msg = str(e)
        error_trace = traceback.format_exc()
        
        return {
            "success": False,
            "message": f"Search operation failed: {error_msg}",
            "error": error_msg,
            "traceback": error_trace
        }

if __name__ == "__main__":
    # Example usage and testing
    if len(sys.argv) >= 4:
        url = sys.argv[1]
        access_key = sys.argv[2]
        password = sys.argv[3]
        
        print(f"Testing authentication to: {url}")
        result = authenticate_seeq(url, access_key, password)
        print(json.dumps(result, indent=2))
    else:
        print("Usage: python seeq_auth.py <url> <access_key> <password>")
        print("Example: python seeq_auth.py https://talosenergy.seeq.tech WoUknZw9SWyiI9K0Y3xu0A 0irkn3dDzvfwgjDXzwjqkHEcKb6zde")
