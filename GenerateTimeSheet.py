import argparse
import json
import pandas as pd
from TimeSheetGenerator import TimeSheetGenerator

# Config and standard argument parsing
parser = argparse.ArgumentParser()
parser.add_argument('config',help="Pass the dates as a config json")
args = parser.parse_args()

config_args = None
with open(args.config,'r') as f:
    config_args = json.loads(f.read())
    


