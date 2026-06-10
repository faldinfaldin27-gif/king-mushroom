"""Train a YOLO26 detector on the King Oyster Mushroom dataset.

    python scripts/train_yolo26.py --epochs 100 --imgsz 640 --model yolo26n.pt

Dataset is dense small objects -> try --imgsz 1024 if recall on tiny mushrooms is low.
Trained weights land in runs/detect/<name>/weights/best.pt; we also copy to models/.
"""
import argparse
import shutil
from pathlib import Path

from ultralytics import YOLO

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data.yaml"
MODELS = ROOT / "models"


def parse_args():
    p = argparse.ArgumentParser(description="Train YOLO26 on king oyster mushrooms")
    p.add_argument("--model", default="yolo26n.pt", help="base model (yolo26n/s/m/l/x.pt)")
    p.add_argument("--data", default=str(DATA))
    p.add_argument("--epochs", type=int, default=100)
    p.add_argument("--imgsz", type=int, default=640)
    p.add_argument("--batch", type=int, default=16)
    p.add_argument("--device", default=None, help="e.g. 0 for GPU, cpu for CPU")
    p.add_argument("--name", default="king_oyster_yolo26")
    return p.parse_args()


def main():
    args = parse_args()
    model = YOLO(args.model)
    results = model.train(
        data=args.data,
        epochs=args.epochs,
        imgsz=args.imgsz,
        batch=args.batch,
        device=args.device,
        name=args.name,
    )
    # Copy best weights to models/ for the other scripts to find.
    best = Path(results.save_dir) / "weights" / "best.pt"
    if best.exists():
        MODELS.mkdir(exist_ok=True)
        dest = MODELS / "best.pt"
        shutil.copy(best, dest)
        print("Best weights copied to:", dest)


if __name__ == "__main__":
    main()
