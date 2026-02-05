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
    # Compute MD5 hash of the date and Dow value
    h = hashlib.md5(datedow, usedforsecurity=False).hexdigest()
    
    # Split hash into two parts for latitude and longitude
    lat_hash = h[:16]
    lon_hash = h[16:]
    
    # Convert hexadecimal to decimal fractions
    lat_offset = int(lat_hash, 16) / (16 ** 16)
    lon_offset = int(lon_hash, 16) / (16 ** 16)
    
    return lat_offset, lon_offset


# Automatically fly when module is imported
fly()
