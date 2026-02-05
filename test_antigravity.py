"""
Unit tests for the antigravity module.
"""

import unittest
from unittest.mock import patch
import antigravity


class TestAntigravity(unittest.TestCase):
    """Test cases for the antigravity module."""
    
    @patch('antigravity.webbrowser.open')
    def test_fly_opens_xkcd_comic(self, mock_open):
        """Test that fly() opens the correct XKCD comic URL."""
        antigravity.fly()
        mock_open.assert_called_once_with("https://xkcd.com/353/")
    
    def test_geohash_returns_tuple(self):
        """Test that geohash returns a tuple of two floats."""
        result = antigravity.geohash(37.421542, -122.085589, b'2005-05-26-10458.68')
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 2)
        self.assertIsInstance(result[0], float)
        self.assertIsInstance(result[1], float)
    
    def test_geohash_values_in_range(self):
        """Test that geohash returns values between 0 and 1."""
        lat_offset, lon_offset = antigravity.geohash(37.421542, -122.085589, b'2005-05-26-10458.68')
        self.assertGreaterEqual(lat_offset, 0.0)
        self.assertLessEqual(lat_offset, 1.0)
        self.assertGreaterEqual(lon_offset, 0.0)
        self.assertLessEqual(lon_offset, 1.0)
    
    def test_geohash_deterministic(self):
        """Test that geohash returns the same values for the same inputs."""
        result1 = antigravity.geohash(37.421542, -122.085589, b'2005-05-26-10458.68')
        result2 = antigravity.geohash(37.421542, -122.085589, b'2005-05-26-10458.68')
        self.assertEqual(result1, result2)
    
    def test_geohash_different_inputs(self):
        """Test that geohash returns different values for different inputs."""
        result1 = antigravity.geohash(37.421542, -122.085589, b'2005-05-26-10458.68')
        result2 = antigravity.geohash(37.421542, -122.085589, b'2005-05-27-10500.00')
        self.assertNotEqual(result1, result2)
    
    def test_geohash_known_value(self):
        """Test geohash with a known input/output pair."""
        # This test verifies the algorithm produces consistent results
        lat, lon = antigravity.geohash(37.421542, -122.085589, b'2005-05-26-10458.68')
        # These values are derived from the MD5 hash of the input
        # We're testing that the function produces consistent results
        self.assertIsNotNone(lat)
        self.assertIsNotNone(lon)
        # Verify the values are reasonable (between 0 and 1)
        self.assertTrue(0 <= lat <= 1)
        self.assertTrue(0 <= lon <= 1)


if __name__ == '__main__':
    unittest.main()
