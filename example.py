#!/usr/bin/env python3
"""
Example usage of the antigravity module.

This script demonstrates the various features of the antigravity module.
"""

from unittest.mock import patch
import antigravity


def main():
    print("=" * 60)
    print("Antigravity Module Examples")
    print("=" * 60)
    print()
    
    # Example 1: Import the module (mocked to prevent browser opening)
    print("Example 1: Importing antigravity module")
    print("-" * 60)
    print("Code: import antigravity")
    print()
    print("✓ Module imported successfully (no auto-fly since Python 3.6+)")
    print("  Note: fly() is no longer called automatically on import.")
    print("  This prevents unwanted side effects in production code.")
    print()
    
    # Example 2: Explicitly call fly()
    print("Example 2: Using the fly() function")
    print("-" * 60)
    print("Code: antigravity.fly()")
    print()
    with patch('antigravity.webbrowser.open') as mock_open:
        mock_open.return_value = True
        result = antigravity.fly()
        if mock_open.called:
            print(f"✓ Would open browser to: {mock_open.call_args[0][0]}")
            print(f"  Function returned: {result}")
        else:
            print("✗ Function did not call webbrowser.open()")
    print()
    
    # Example 3: Use geohash function
    print("Example 3: Computing a geohash")
    print("-" * 60)
    print("Code:")
    print("  latitude = 37.421542")
    print("  longitude = -122.085589")
    print("  date_dow = b'2005-05-26-10458.68'")
    print("  lat_offset, lon_offset = antigravity.geohash(latitude, longitude, date_dow)")
    print()
    
    latitude = 37.421542
    longitude = -122.085589
    date_dow = b'2005-05-26-10458.68'
    
    try:
        lat_offset, lon_offset = antigravity.geohash(latitude, longitude, date_dow)
        
        print(f"Result:")
        print(f"  Latitude offset:  {lat_offset:.10f}")
        print(f"  Longitude offset: {lon_offset:.10f}")
        print(f"  New latitude:     {latitude + lat_offset:.6f}")
        print(f"  New longitude:    {longitude + lon_offset:.6f}")
    except (TypeError, ValueError) as e:
        print(f"✗ Error: {e}")
    print()
    
    # Example 4: Different date produces different hash
    print("Example 4: Geohash with different date")
    print("-" * 60)
    date_dow2 = b'2005-05-27-10500.00'
    try:
        lat_offset2, lon_offset2 = antigravity.geohash(latitude, longitude, date_dow2)
        print(f"With date_dow = {date_dow2.decode('utf-8')}:")
        print(f"  Latitude offset:  {lat_offset2:.10f}")
        print(f"  Longitude offset: {lon_offset2:.10f}")
    except (TypeError, ValueError, UnicodeDecodeError) as e:
        print(f"✗ Error: {e}")
    print()
    
    print("=" * 60)
    print("All examples completed successfully!")
    print("=" * 60)


if __name__ == '__main__':
    main()
