# Antigravity

A Python implementation of the famous XKCD "antigravity" Easter egg, inspired by [XKCD comic #353](https://xkcd.com/353/).

## What is this?

In the XKCD comic, a character demonstrates the power and simplicity of Python by typing `import antigravity`, which causes them to fly. This repository implements that joke as a real Python module.

## Features

- **Automatic flying**: Simply importing the module opens the XKCD Python comic in your default web browser
- **Geohashing**: Includes the XKCD geohashing algorithm from [XKCD comic #426](https://xkcd.com/426/)
- **Pure Python**: No external dependencies required

## Installation

You can install this package directly from the repository:

```bash
pip install git+https://github.com/ABCPA/Antigravity.git
```

Or clone and install locally:

```bash
git clone https://github.com/ABCPA/Antigravity.git
cd Antigravity
pip install -e .
```

## Usage

### Flying (Opening XKCD Comic)

Simply import the module to open the XKCD Python comic:

```python
import antigravity
```

Or call the `fly()` function explicitly:

```python
from antigravity import fly
fly()
```

### Geohashing

Use the geohashing algorithm to compute location offsets:

```python
from antigravity import geohash

# Compute geohash for a specific date and location
latitude = 37.421542
longitude = -122.085589
date_dow = b'2005-05-26-10458.68'

lat_offset, lon_offset = geohash(latitude, longitude, date_dow)
print(f"Latitude offset: {lat_offset}")
print(f"Longitude offset: {lon_offset}")
```

## How It Works

When you import the `antigravity` module, it automatically calls the `fly()` function, which uses Python's `webbrowser` module to open the XKCD Python comic in your default browser. This is a humorous reference to the comic's premise that Python is so powerful that a single import statement can make you fly.

The geohashing function implements the algorithm described in XKCD #426, which uses MD5 hashing to generate pseudo-random but deterministic location offsets based on a date and the Dow Jones opening value.

## About XKCD #353

The comic shows a character floating in the air, explaining to friends:

> "I wrote 20 short programs in Python yesterday. It was wonderful. Perl, I'm leaving you."
> 
> "I learned it last night! Everything is so simple! Hello world is just 'print "Hello, World!"'"
> 
> "I'm flying!"
> 
> "How are you flying?"
> 
> "Python!"
> 
> "I just typed 'import antigravity'"

This captures the joy and simplicity that many programmers feel when discovering Python.

## License

MIT License

## Credits

- Inspired by [XKCD comic #353](https://xkcd.com/353/) by Randall Munroe
- Geohashing algorithm from [XKCD comic #426](https://xkcd.com/426/)
- Original Python antigravity module in the standard library