# Hash Excel Columns

Load Excel file, choose a column and confirm to hash this one.

These dependencies are used:

- PySide2
- Fbs
- OpenPyXl
- Hashlib

See [requirements](requirements.txt) for details

## System Requirements

Qt for Python is free until Python 3.6


## Installation

1. Create venv dir

````bash
python -m venv venv
````

2. Activate venv

````bash
.\venv\Scripts\activate
````

3. Install dependencies

````bash
pip install -r requirements.txt
````

## Run Application

````bash
fbs run
````

## Create Executable

````bash
fbs freeze
````

Then


````bash
fbs installer
````
