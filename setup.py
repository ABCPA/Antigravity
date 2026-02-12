"""
Setup script for the antigravity package.
"""

from setuptools import setup, find_packages
import os

# Read long description from README with error handling
try:
    with open("README.md", "r", encoding="utf-8") as fh:
        long_description = fh.read()
except FileNotFoundError:
    long_description = "A Python Easter egg inspired by XKCD comic #353"

setup(
    name="antigravity",
    version="1.0.0",
    author="ABCPA",
    author_email="noreply@github.com",
    description="A Python Easter egg inspired by XKCD comic #353",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/ABCPA/Antigravity",
    py_modules=["antigravity"],
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
    ],
    python_requires=">=3.7",
    install_requires=[
        # No external dependencies - uses only Python standard library
    ],
    license="MIT",
    license_files=["LICENSE"],
)
