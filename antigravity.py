"""
Antigravity module - A Python Easter egg inspired by XKCD comic #353.

This module implements the famous "antigravity" joke from XKCD.
When imported, it opens the XKCD Python comic in your default web browser.
"""

import webbrowser
import hashlib


def fly():
    """
    Open the XKCD Python comic (https://xkcd.com/353/) in the default web browser.
    
    This is the main antigravity function that makes Python programmers "fly"
    by demonstrating the simplicity and power of Python.
    """
    webbrowser.open("https://xkcd.com/353/")


def geohash(latitude, longitude, datedow):
    """
    Compute geohash using the XKCD algorithm.
    
    This implements the geohashing algorithm from XKCD comic #426.
    The algorithm uses the integer parts of latitude/longitude to define
    a 1°×1° graticule, then hashes the date+DJIA to get pseudo-random
    fractional offsets within that square.
    
    Args:
        latitude (float): The latitude coordinate
        longitude (float): The longitude coordinate  
        datedow (bytes): The date and Dow Jones opening value
        
    Returns:
        tuple: A tuple of (latitude_offset, longitude_offset) between 0 and 1
        
    Example:
        >>> lat, lon = geohash(37.421542, -122.085589, b'2005-05-26-10458.68')
        >>> print(f"Offset: {lat:.6f}, {lon:.6f}")
    """
    # Get integer parts for the graticule (1°×1° square)
    lat_int = int(latitude)
    lon_int = int(longitude)
    
    # Create the hash input by combining graticule and date info
    # This ensures different graticules get different results for the same date
    hash_input = f"{lat_int},{lon_int},{datedow.decode()}".encode()
    
    # Compute MD5 hash - compatible with Python 3.6+
    h = hashlib.md5(hash_input).hexdigest()
    
    # Split hash into two parts for latitude and longitude
    lat_hash = h[:16]
    lon_hash = h[16:]
    
    # Convert hexadecimal to decimal fractions
    lat_offset = int(lat_hash, 16) / (16 ** 16)
    lon_offset = int(lon_hash, 16) / (16 ** 16)
    
    return lat_offset, lon_offset


# Automatically fly when module is imported
fly()
