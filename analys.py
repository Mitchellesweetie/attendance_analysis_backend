import pandas as pd
from pathlib import Path

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

Path(OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)
Path(UPLOAD_FOLDER).mkdir(parents=True, exist_ok=True)

def processdata():
    return 'hello'
