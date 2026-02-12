"""
Antigravity module - A Python Easter egg inspired by XKCD comic #353.

This module implements the famous "antigravity" joke from XKCD.
When imported, it opens the XKCD Python comic in your default web browser.
"""

import webbrowser
import hashlib
from typing import Tuple


def fly() -> bool:
    """
    Open the XKCD Python comic (https://xkcd.com/353/) in the default web browser.
    
    This is the main antigravity function that makes Python programmers "fly"
    by demonstrating the simplicity and power of Python.
    
    Returns:
        bool: True if browser was opened successfully, False otherwise
    """
    try:
        return webbrowser.open("https://xkcd.com/353/")
    except Exception as e:
        # Log or print error if needed, but don't crash
        print(f"Warning: Could not open browser: {e}")
        return False


def geohash(latitude: float, longitude: float, datedow: bytes) -> Tuple[float, float]:
    """
    Compute geohash using the XKCD algorithm.
    
    This implements the geohashing algorithm from XKCD comic #426.
    
    Args:
        latitude (float): The latitude coordinate (not used in hash computation,
                         but required for computing final position)
        longitude (float): The longitude coordinate (not used in hash computation,
                          but required for computing final position)
        datedow (bytes): The date and Dow Jones opening value used for hash computation
        
    Returns:
        tuple: A tuple of (latitude_offset, longitude_offset) between 0 and 1
        
    Raises:
        TypeError: If datedow is not bytes
        ValueError: If datedow is empty
        
    Example:
        >>> lat, lon = geohash(37.421542, -122.085589, b'2005-05-26-10458.68')
        >>> print(f"Offset: {lat:.6f}, {lon:.6f}")
    
    Note:
        The latitude and longitude parameters are included for API compatibility
        but are not used in the hash computation. Only datedow affects the result.
        To get the final coordinates, add the offsets to the input coordinates:
        final_lat = latitude + lat_offset
        final_lon = longitude + lon_offset
    """
    # Input validation
    if not isinstance(datedow, bytes):
        raise TypeError("datedow must be bytes")
    if not datedow:
        raise ValueError("datedow cannot be empty")
    
    # Compute MD5 hash of the date and Dow value
    h = hashlib.md5(datedow, usedforsecurity=False).hexdigest()
    
    # Split hash into two parts for latitude and longitude
    lat_hash = h[:16]
    lon_hash = h[16:]
    
    # Convert hexadecimal to decimal fractions
    lat_offset = int(lat_hash, 16) / (16 ** 16)
    lon_offset = int(lon_hash, 16) / (16 ** 16)
    
    return lat_offset, lon_offset


# Module can be imported without side effects
# Call fly() explicitly if you want to open the browser
if __name__ == '__main__':
    fly()
