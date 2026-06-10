"""Detect and count king oyster mushrooms in a bag photo (tasks 1 & 2).

    python scripts/detect.py --weights models/best.pt --source img.jpg --count
    python scripts/detect.py --source dataset_roboflow/test/images --count

Saves annotated images to outputs/ and prints a per-image count.
Falls back to a pretrained YOLO26 checkpoint if no trained weights are given
(it won't actually know mushrooms until you train -- see train_yolo26.py).
"""
import argparse
from pathlib import Path

from ultralytics import YOLO

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs"


def parse_args():
    p = argparse.ArgumentParser(description="Detect + count king oyster mushrooms")
    p.add_argument("--weights", default=str(ROOT / "models" / "best.pt"))
    p.add_argument("--source", required=True, help="image, folder, or video")
    p.add_argument("--conf", type=float, default=0.25)
    p.add_argument("--imgsz", type=int, default=640)
    p.add_argument("--count", action="store_true", help="print mushroom count per image")
    return p.parse_args()


def main():
    args = parse_args()
    weights = args.weights if Path(args.weights).exists() else "yolo26n.pt"
    if weights != args.weights:
        print(f"[warn] {args.weights} not found -> using pretrained {weights} "
              f"(train first for real mushroom detection).")

    model = YOLO(weights)
    OUT.mkdir(exist_ok=True)
    results = model.predict(
        source=args.source, conf=args.conf, imgsz=args.imgsz,
        save=True, project=str(OUT), name="detect", exist_ok=True,
    )

    total = 0
    for r in results:
        n = len(r.boxes)
        total += n
        if args.count:
            print(f"{Path(r.path).name}: {n} mushrooms")
    print(f"\nTotal across {len(results)} image(s): {total} mushrooms")
    print(f"Annotated images saved to: {OUT / 'detect'}")


if __name__ == "__main__":
    main()
