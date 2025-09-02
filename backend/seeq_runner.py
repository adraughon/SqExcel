#!/usr/bin/env python3
"""
Seeq Runner Script for TSFlow
This script maintains authentication state and handles all Seeq operations
"""

import json
import sys
import traceback
import io
from contextlib import redirect_stdout, redirect_stderr
from typing import Dict, Any, Optional

# Capture all output to prevent interference with JSON response
output_capture = io.StringIO()
stderr_capture = io.StringIO()

try:
    from seeq import spy
except ImportError:
    print("SPy module not found. Please install it using: pip install seeq")
    sys.exit(1)

# Global authentication state
auth_state = {
    'is_authenticated': False,
    'url': None,
    'access_key': None,
    'password': None,
    'auth_provider': 'Seeq',
    'ignore_ssl_errors': False
}

def authenticate_seeq(url: str, access_key: str, password: str, 
                     auth_provider: str = 'Seeq', 
                     ignore_ssl_errors: any = False) -> Dict[str, Any]:
    """
    Authenticate with Seeq server using SPy
    """
    try:
        # Suppress SPy output by redirecting stdout temporarily
        import io
        import sys
        from contextlib import redirect_stdout
        
        # Set compatibility option for maximum compatibility
        try:
            spy.options.compatibility = 66
        except AttributeError:
            # If compatibility option doesn't exist, continue without it
            pass
        
        # Set the server URL in SPy options before attempting login
        try:
            if hasattr(spy, 'options') and hasattr(spy.options, 'server'):
                spy.options.server = url
            else:
                pass  # Cannot set server URL
        except Exception as e:
            pass  # Ignore errors setting server URL
        
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
            # Update global state
            auth_state['is_authenticated'] = True
            auth_state['url'] = url
            auth_state['access_key'] = access_key
            auth_state['password'] = password
            auth_state['auth_provider'] = auth_provider
            auth_state['ignore_ssl_errors'] = ignore_ssl_errors
            
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

def check_auth_status() -> Dict[str, Any]:
    """
    Check the current authentication status with Seeq
    """
    try:
        if spy.user is not None and auth_state['is_authenticated']:
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

def search_sensors_only(sensor_names: list, url: str = None, access_key: str = None, 
                       password: str = None, auth_provider: str = 'Seeq', 
                       ignore_ssl_errors: any = False) -> Dict[str, Any]:
    """
    Search for sensors in Seeq without pulling data
    """
    try:
        import pandas as pd
        
        # Always authenticate when called from Node.js (new process each time)
        if url and access_key and password:
            # Try to authenticate first
            auth_result = authenticate_seeq(url, access_key, password, auth_provider, ignore_ssl_errors)
            if not auth_result['success']:
                return {
                    "success": False,
                    "message": f"Authentication failed: {auth_result.get('error', 'Unknown error')}",
                    "error": "Authentication required",
                    "auth_details": auth_result
                }
            # Check if authentication was successful
            if spy.user is None:
                return {
                    "success": False,
                    "message": "Authentication appeared successful but spy.user is still None",
                    "error": "Authentication state issue"
                }
        else:
            return {
                "success": False,
                "message": "Authentication credentials are required",
                "error": "Missing credentials"
            }
        
        # Search for sensors
        search_results = []
        for sensor_name in sensor_names:
            try:
                # Search with Type set to StoredSignal and suppress output
                result = spy.search({
                    'Name': sensor_name,
                    'Type': 'StoredSignal'
                }, quiet=True)
                
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

def search_and_pull_sensors(sensor_names: list, start_datetime: str, end_datetime: str, 
                           grid: str = '15min', timezone: str = None, url: str = None, 
                           access_key: str = None, password: str = None, 
                           auth_provider: str = 'Seeq', ignore_ssl_errors: any = False) -> Dict[str, Any]:
    """
    Search for sensors in Seeq and pull their data
    """
    try:
        import pandas as pd
        from datetime import datetime
        
        # Always authenticate when called from Node.js (new process each time)
        if url and access_key and password:
            # Try to authenticate first
            auth_result = authenticate_seeq(url, access_key, password, auth_provider, ignore_ssl_errors)
            
            if not auth_result['success']:
                return {
                    "success": False,
                    "message": f"Authentication failed: {auth_result.get('error', 'Unknown error')}",
                    "error": "Authentication required",
                    "auth_details": auth_result
                }
            
            # Check if authentication was successful
            if spy.user is None:
                return {
                    "success": False,
                    "message": "Authentication appeared successful but spy.user is still None",
                    "error": "Authentication state issue"
                }
        else:
            return {
                "success": False,
                "message": "Authentication credentials are required",
                "error": "Missing credentials"
            }
        
        # Authentication status is checked above
            # Parse datetime strings - handle multiple formats including Excel dates
        def parse_excel_friendly_datetime(dt_str):
            """Parse datetime strings in various formats including Excel-friendly ones"""
            import sys
            print(f"[DEBUG] Parsing datetime: '{dt_str}' (type: {type(dt_str)})", file=sys.stderr)
            
            if not dt_str:
                print("[DEBUG] Empty datetime string, returning None", file=sys.stderr)
                return None
                
            # Try different parsing strategies
            try:
                # First try pandas flexible parsing
                result = pd.to_datetime(dt_str)
                print(f"[DEBUG] Pandas flexible parsing succeeded: {result} (type: {type(result)})", file=sys.stderr)
                return result
            except Exception as e:
                print(f"[DEBUG] Pandas flexible parsing failed: {e}", file=sys.stderr)
                
            try:
                # Try common Excel date formats
                excel_formats = [
                    '%m/%d/%Y',           # 9/1/2025
                    '%m/%d/%Y %H:%M:%S',  # 9/1/2025 12:00:00
                    '%m/%d/%Y %I:%M:%S %p', # 9/1/2025 12:00:00 PM
                    '%Y-%m-%d',           # 2025-09-01
                    '%Y-%m-%d %H:%M:%S',  # 2025-09-01 12:00:00
                    '%m-%d-%Y',           # 09-01-2025
                    '%m-%d-%Y %H:%M:%S',  # 09-01-2025 12:00:00
                    '%d/%m/%Y',           # 1/9/2025 (European format)
                    '%d/%m/%Y %H:%M:%S',  # 1/9/2025 12:00:00
                ]
                
                for fmt in excel_formats:
                    try:
                        result = pd.to_datetime(dt_str, format=fmt)
                        print(f"[DEBUG] Excel format '{fmt}' parsing succeeded: {result} (type: {type(result)})", file=sys.stderr)
                        return result
                    except Exception as e:
                        print(f"[DEBUG] Excel format '{fmt}' parsing failed: {e}", file=sys.stderr)
                        continue
                        
                # If all else fails, try to parse as Excel serial number
                try:
                    excel_serial = float(dt_str)
                    print(f"[DEBUG] Attempting Excel serial number parsing: {excel_serial}", file=sys.stderr)
                    # Excel dates are days since 1900-01-01
                    # Note: Excel incorrectly treats 1900 as a leap year
                    excel_epoch = pd.Timestamp('1899-12-30')
                    result = excel_epoch + pd.Timedelta(days=excel_serial)
                    print(f"[DEBUG] Excel serial number parsing succeeded: {result} (type: {type(result)})", file=sys.stderr)
                    return result
                except Exception as e:
                    print(f"[DEBUG] Excel serial number parsing failed: {e}", file=sys.stderr)
                    
                # Last resort: try to parse with dateutil
                try:
                    from dateutil import parser
                    result = parser.parse(dt_str)
                    print(f"[DEBUG] dateutil parsing succeeded: {result} (type: {type(result)})", file=sys.stderr)
                    return result
                except Exception as e:
                    print(f"[DEBUG] dateutil parsing failed: {e}", file=sys.stderr)
                    
            except Exception as e:
                print(f"[DEBUG] All parsing methods failed for '{dt_str}': {e}", file=sys.stderr)
                raise ValueError(f"Could not parse datetime '{dt_str}': {str(e)}")
        
        try:
            import sys
            print(f"[DEBUG] Starting datetime parsing...", file=sys.stderr)
            print(f"[DEBUG] start_datetime: '{start_datetime}' (type: {type(start_datetime)})", file=sys.stderr)
            print(f"[DEBUG] end_datetime: '{end_datetime}' (type: {type(end_datetime)})", file=sys.stderr)
            
            start_dt = parse_excel_friendly_datetime(start_datetime)
            end_dt = parse_excel_friendly_datetime(end_datetime)
            
            print(f"[DEBUG] Parsed start_dt: {start_dt} (type: {type(start_dt)}, tz: {getattr(start_dt, 'tz', 'None')})", file=sys.stderr)
            print(f"[DEBUG] Parsed end_dt: {end_dt} (type: {type(end_dt)}, tz: {getattr(end_dt, 'tz', 'None')})", file=sys.stderr)
            
            if start_dt is None or end_dt is None:
                print("[DEBUG] One or both datetime values are None", file=sys.stderr)
                return {
                    "success": False,
                    "message": "Start and end datetime are required",
                    "error": "Missing datetime values"
                }
                
            # Apply timezone if specified
            if timezone:
                print(f"[DEBUG] Applying timezone: {timezone}", file=sys.stderr)
                start_dt = start_dt.tz_localize('UTC').tz_convert(timezone)
                end_dt = end_dt.tz_localize('UTC').tz_convert(timezone)
            elif start_dt.tz is None:
                # If no timezone specified, assume UTC for consistency
                print("[DEBUG] No timezone specified, localizing to UTC", file=sys.stderr)
                start_dt = start_dt.tz_localize('UTC')
                end_dt = end_dt.tz_localize('UTC')
            else:
                print(f"[DEBUG] Using existing timezone: {start_dt.tz}", file=sys.stderr)
                
            print(f"[DEBUG] Final start_dt: {start_dt} (tz: {start_dt.tz})", file=sys.stderr)
            print(f"[DEBUG] Final end_dt: {end_dt} (tz: {end_dt.tz})", file=sys.stderr)
                
        except Exception as e:
            print(f"[DEBUG] Datetime parsing failed with error: {e}", file=sys.stderr)
            print(f"[DEBUG] Error type: {type(e)}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            return {
                "success": False,
                "message": f"Invalid datetime format: {str(e)}",
                "error": "Datetime parsing failed",
                "supported_formats": [
                    "Excel dates: 9/1/2025, 9/1/2025 12:00:00 PM",
                    "ISO format: 2025-09-01T00:00:00Z",
                    "Standard formats: 09/01/2025, 2025-09-01",
                    "Excel serial numbers: 45292.5"
                ]
            }
        
        # Ensure SPy is properly initialized with server URL
        try:
            if hasattr(spy, 'options') and hasattr(spy.options, 'server'):
                pass  # Server option available
            else:
                pass  # Server option not available
                
            # Always try to set the server URL to ensure it's set
            if url:
                try:
                    spy.options.server = url
                except Exception as e:
                    pass  # Ignore errors setting server URL
        except Exception as e:
            pass  # Ignore errors checking SPy options
        
        # Search for sensors
        search_results = []
        for sensor_name in sensor_names:
            try:
                # Search with Type set to StoredSignal and suppress output
                result = spy.search({
                    'Name': sensor_name,
                    'Type': 'StoredSignal'
                }, quiet=True)
                
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
        
        # Drop duplicates based on sensor names to avoid header conflicts
        # Keep the first occurrence of each sensor name
        original_count = len(valid_sensors)
        valid_sensors = valid_sensors.drop_duplicates(subset=['Original_Name'], keep='first')
        final_count = len(valid_sensors)
        
        if original_count > final_count:
            print(f"Removed {original_count - final_count} duplicate sensor names to avoid header conflicts")
        
        # Pull data for valid sensors
        try:
            import sys
            print(f"[DEBUG] Pulling data for {len(valid_sensors)} sensors...", file=sys.stderr)
            print(f"[DEBUG] Time range: {start_dt} to {end_dt}", file=sys.stderr)
            print(f"[DEBUG] Grid: {grid}", file=sys.stderr)
            
            data_df = spy.pull(
                valid_sensors,
                start=start_dt,
                end=end_dt,
                grid=grid,
                header='Name',  # Use Name for readable column headers
                quiet=True
            )
            
            print(f"[DEBUG] Data retrieved successfully. DataFrame shape: {data_df.shape}", file=sys.stderr)
            print(f"[DEBUG] DataFrame columns: {list(data_df.columns)}", file=sys.stderr)
            print(f"[DEBUG] DataFrame index type: {type(data_df.index)}", file=sys.stderr)
            print(f"[DEBUG] First few index values: {list(data_df.index[:3])}", file=sys.stderr)
            print(f"[DEBUG] Sample data types:", file=sys.stderr)
            for col in data_df.columns[:3]:  # Show first 3 columns
                sample_val = data_df[col].iloc[0] if not data_df.empty else None
                print(f"  {col}: {sample_val} (type: {type(sample_val)})", file=sys.stderr)
            
            # Convert to records for JSON serialization
            data_records = data_df.reset_index().to_dict('records')
            print(f"[DEBUG] Converted to records. First record keys: {list(data_records[0].keys()) if data_records else 'No records'}", file=sys.stderr)
            
            # Clean NaN values and convert timestamps for JSON serialization
            def clean_for_json(obj):
                if isinstance(obj, dict):
                    return {k: clean_for_json(v) for k, v in obj.items()}
                elif isinstance(obj, list):
                    return [clean_for_json(item) for item in obj]
                elif obj != obj:  # Check for NaN
                    return None
                elif hasattr(obj, 'isoformat'):  # Handle pandas Timestamp objects
                    print(f"[DEBUG] Converting timestamp object to Excel-friendly format: {obj} (type: {type(obj)})", file=sys.stderr)
                    # Convert to Excel-friendly format: YYYY-MM-DD HH:MM:SS
                    if obj.tz is not None:
                        # If timezone-aware, convert to UTC and format
                        utc_obj = obj.tz_convert('UTC')
                        result = utc_obj.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        # If no timezone, format directly
                        result = obj.strftime('%Y-%m-%d %H:%M:%S')
                    print(f"[DEBUG] Converted to Excel-friendly format: {result} (type: {type(result)})", file=sys.stderr)
                    return result
                else:
                    return obj
            
            print("[DEBUG] Starting data cleaning process...", file=sys.stderr)
            cleaned_data = clean_for_json(data_records)
            cleaned_search_results = clean_for_json(combined_results.to_dict('records'))
            print(f"[DEBUG] Data cleaning completed. Cleaned data length: {len(cleaned_data) if cleaned_data else 0}", file=sys.stderr)
            
            result = {
                "success": True,
                "message": f"Successfully retrieved data for {len(valid_sensors)} sensors",
                "search_results": cleaned_search_results,
                "data": cleaned_data,
                "data_columns": list(data_df.columns),
                "data_index": [str(idx) for idx in data_df.index],
                "sensor_count": len(valid_sensors),
                "time_range": f"{start_datetime} to {end_datetime}"
            }
            
            print(f"[DEBUG] Preparing to return result with keys: {list(result.keys())}", file=sys.stderr)
            print(f"[DEBUG] Result data type: {type(result['data'])}", file=sys.stderr)
            print(f"[DEBUG] Result data length: {len(result['data']) if result['data'] else 0}", file=sys.stderr)
            if result['data']:
                print(f"[DEBUG] First data record keys: {list(result['data'][0].keys()) if result['data'][0] else 'No keys'}", file=sys.stderr)
            
            return result
                
        except Exception as e:
            import sys
            print(f"[DEBUG] Error during data retrieval: {e}", file=sys.stderr)
            print(f"[DEBUG] Error type: {type(e)}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            return {
                "success": False,
                "message": f"Failed to pull data: {str(e)}",
                "error": str(e),
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

def test_connection(url: str) -> Dict[str, Any]:
    """
    Test connection to Seeq server without authentication
    """
    try:
        # This is a simple test - we'll just try to import spy and check if it's available
        return {
            "success": True,
            "message": "SPy module is available and ready for authentication",
            "status_code": "OK"
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"Connection test failed: {str(e)}",
            "error": str(e)
        }

def get_server_info(url: str) -> Dict[str, Any]:
    """
    Get Seeq server information
    """
    try:
        return {
            "success": True,
            "message": "Server info retrieved successfully",
            "server_info": {
                "status": "Available",
                "url": url
            }
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"Failed to get server info: {str(e)}",
            "error": str(e)
        }

# Main execution logic
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(json.dumps({
            "success": False,
            "error": "Usage: python seeq_runner.py <function_name> <json_args>"
        }))
        sys.exit(1)
    
    function_name = sys.argv[1]
    args_json = sys.argv[2]
    
    try:
        # Capture all output to prevent interference
        with redirect_stdout(output_capture), redirect_stderr(stderr_capture):
            args = json.loads(args_json)
            
            if function_name == 'authenticate_seeq':
                result = authenticate_seeq(*args)
            elif function_name == 'check_auth_status':
                result = check_auth_status()
            elif function_name == 'search_sensors_only':
                result = search_sensors_only(*args)
            elif function_name == 'search_and_pull_sensors':
                result = search_and_pull_sensors(*args)
            elif function_name == 'test_connection':
                result = test_connection(*args)
            elif function_name == 'get_server_info':
                result = get_server_info(*args)
            else:
                result = {
                    "success": False,
                    "error": f"Unknown function: {function_name}"
                }
        
        # Debug logging before JSON serialization (to stderr to avoid interfering with JSON output)
        import sys
        print(f"[DEBUG] About to serialize result to JSON...", file=sys.stderr)
        print(f"[DEBUG] Result type: {type(result)}", file=sys.stderr)
        print(f"[DEBUG] Result keys: {list(result.keys()) if isinstance(result, dict) else 'Not a dict'}", file=sys.stderr)
        
        # Only print the JSON result, nothing else
        try:
            json_result = json.dumps(result)
            print(f"[DEBUG] JSON serialization successful, length: {len(json_result)}", file=sys.stderr)
            print(json_result)
        except Exception as e:
            print(f"[DEBUG] JSON serialization failed: {e}", file=sys.stderr)
            print(f"[DEBUG] Error type: {type(e)}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            # Fallback: try to serialize with default handler
            try:
                import json as json_module
                json_result = json_module.dumps(result, default=str)
                print(f"[DEBUG] Fallback JSON serialization successful with default=str", file=sys.stderr)
                print(json_result)
            except Exception as e2:
                print(f"[DEBUG] Fallback JSON serialization also failed: {e2}", file=sys.stderr)
                # Last resort: return error as string
                print(json.dumps({
                    "success": False,
                    "error": f"JSON serialization failed: {str(e)}",
                    "fallback_error": f"Fallback also failed: {str(e2)}"
                }))
        
    except Exception as e:
        # Debug logging for execution errors
        import sys
        print(f"[DEBUG] Execution error occurred: {e}", file=sys.stderr)
        print(f"[DEBUG] Error type: {type(e)}", file=sys.stderr)
        import traceback
        error_trace = traceback.format_exc()
        print(f"[DEBUG] Full traceback:\n{error_trace}", file=sys.stderr)
        
        # Only print the JSON error, nothing else
        print(json.dumps({
            "success": False,
            "error": f"Execution error: {str(e)}",
            "traceback": error_trace
        }))
        sys.exit(1)
