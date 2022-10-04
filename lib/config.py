import argparse
import json
import os
from types import SimpleNamespace

args = argparse.ArgumentParser()
args.add_argument('--config', default=False,
                  action=argparse.BooleanOptionalAction)
args = args.parse_args()

pwd = os.getcwd()

def getConfig():
    if args.config or "excel-to-pdf.json" not in os.listdir():
        conf = {
            "input": "doc.xlsx",
            "output": "build/doc.pdf",
            "logo": "logos/logo.png",
            "config": "excel-to-pdf.json",
        }
        print("Creating config file...")
        with open("excel-to-pdf.json", "w") as f:
            f.write(json.dumps(conf))
    else:
        print("Loading config file...")
        with open("excel-to-pdf.json", "r") as f:
            conf = json.loads(f.read(), object_hook=lambda d: SimpleNamespace(**d))
        # replace each conf key with the absolute path
        for key in conf.__dict__.keys():
            conf.__dict__[key] = os.path.join(pwd, conf.__dict__[key])

    return conf
