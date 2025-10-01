from pathlib import Path
import pandas as pd
ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
def load_data(filename="NPS_dataset.xlsx"):
    path = DATA_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"File not found at {path}")
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [col.strip() for col in df.columns]
    return df
def save_to_excel(writer, df, sheet_name):
    df.to_excel(writer, sheet_name=sheet_name, index=False)
