"""Measure each mushroom's size (task 3): YOLO26 boxes -> SAM masks -> mm.

    python scripts/measure_size.py --weights models/best.pt --source img.jpg \
        --sam-checkpoint sam_vit_h.pth --bag-width-mm 120

Pipeline:
  1. YOLO26 detects each mushroom (box).
  2. SAM refines each box into a precise mask (optional; falls back to the box).
  3. Convert pixel size -> mm using a scale:
       - --bag-width-mm : real diameter of the bag mouth (太空包口徑). We estimate the
         bag width in pixels from the spread of detections and derive px-per-mm.
       - otherwise sizes are reported in PIXELS only.
  4. Writes outputs/sizes.csv with per-mushroom width/height/area.

SAM is optional: without --sam-checkpoint we measure the YOLO box directly, which is
already a usable size estimate for the assignment.
"""
import argparse
import csv
from pathlib import Path

import cv2
import numpy as np
from ultralytics import YOLO

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs"


def parse_args():
    p = argparse.ArgumentParser(description="Measure king oyster mushroom sizes")
    p.add_argument("--weights", default=str(ROOT / "models" / "best.pt"))
    p.add_argument("--source", required=True, help="single image")
    p.add_argument("--conf", type=float, default=0.25)
    p.add_argument("--imgsz", type=int, default=640)
    p.add_argument("--sam-checkpoint", default=None, help="SAM weights (.pth); optional")
    p.add_argument("--sam-type", default="vit_h", help="vit_h | vit_l | vit_b")
    p.add_argument("--bag-width-mm", type=float, default=None,
                   help="real bag mouth width in mm for px->mm scale")
    return p.parse_args()


def load_sam(checkpoint, model_type):
    """Return a SAM predictor, or None if SAM isn't available."""
    if not checkpoint:
        return None
    try:
        from segment_anything import sam_model_registry, SamPredictor
    except ImportError:
        print("[warn] segment-anything not installed -> measuring boxes only.")
        return None
    sam = sam_model_registry[model_type](checkpoint=checkpoint)
    return SamPredictor(sam)


def main():
    args = parse_args()
    weights = args.weights if Path(args.weights).exists() else "yolo26n.pt"
    model = YOLO(weights)
    img = cv2.imread(args.source)
    if img is None:
        raise SystemExit(f"Could not read image: {args.source}")

    res = model.predict(source=args.source, conf=args.conf, imgsz=args.imgsz)[0]
    boxes = res.boxes.xyxy.cpu().numpy() if res.boxes is not None else np.empty((0, 4))

    predictor = load_sam(args.sam_checkpoint, args.sam_type)
    if predictor is not None:
        predictor.set_image(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))

    # px -> mm scale from the bag width, estimated as the horizontal spread of detections.
    px_per_mm = None
    if args.bag_width_mm and len(boxes):
        bag_px = boxes[:, 2].max() - boxes[:, 0].min()
        if bag_px > 0:
            px_per_mm = bag_px / args.bag_width_mm
            print(f"Scale: bag width ~{bag_px:.0f}px = {args.bag_width_mm}mm "
                  f"-> {px_per_mm:.3f} px/mm")

    OUT.mkdir(exist_ok=True)
    unit = "mm" if px_per_mm else "px"
    rows = []
    vis = img.copy()
    for i, (x1, y1, x2, y2) in enumerate(boxes):
        w_px, h_px = x2 - x1, y2 - y1
        area_px = w_px * h_px

        if predictor is not None:
            mask, _, _ = predictor.predict(
                box=np.array([x1, y1, x2, y2]), multimask_output=False)
            area_px = float(mask[0].sum())
            ys, xs = np.where(mask[0])
            if len(xs):
                w_px, h_px = xs.max() - xs.min(), ys.max() - ys.min()

        if px_per_mm:
            w, h = w_px / px_per_mm, h_px / px_per_mm
            area = area_px / (px_per_mm ** 2)
        else:
            w, h, area = w_px, h_px, area_px

        rows.append([i, round(float(w), 2), round(float(h), 2), round(float(area), 2)])
        cv2.rectangle(vis, (int(x1), int(y1)), (int(x2), int(y2)), (0, 255, 0), 1)
        cv2.putText(vis, f"{w:.0f}x{h:.0f}{unit}", (int(x1), int(y1) - 2),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.3, (0, 0, 255), 1)

    csv_path = OUT / "sizes.csv"
    with open(csv_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["id", f"width_{unit}", f"height_{unit}", f"area_{unit}2"])
        writer.writerows(rows)

    vis_path = OUT / f"sizes_{Path(args.source).stem}.jpg"
    cv2.imwrite(str(vis_path), vis)
    print(f"Measured {len(rows)} mushrooms ({unit}). CSV: {csv_path}  image: {vis_path}")


if __name__ == "__main__":
    main()
