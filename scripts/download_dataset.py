"""Download the King Oyster Mushroom dataset from Roboflow (YOLO format).

esdl/king-oyster-mushroom v8 -> ./dataset_roboflow/{train,valid,test}

Tip: set ROBOFLOW_API_KEY as an env var instead of hard-coding the key.
    python scripts/download_dataset.py
"""
import os
from roboflow import Roboflow

API_KEY = os.environ.get("ROBOFLOW_API_KEY", "A8K2N5mHFh8kTO7xwNRV")
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEST = os.path.join(HERE, "dataset_roboflow")


def main():
    rf = Roboflow(api_key=API_KEY)
    project = rf.workspace("esdl").project("king-oyster-mushroom")
    version = project.version(8)
    # "yolov11" export format is plain YOLO txt labels, compatible with YOLO26 training.
    dataset = version.download("yolov11", location=DEST, overwrite=True)
    print("Downloaded to:", dataset.location)


if __name__ == "__main__":
    main()
