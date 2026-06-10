"""Agentic object detection (task 4) — VisionAgent-style LLM/VLM over the image.

    set ANTHROPIC_API_KEY=...
    python scripts/agentic_detect.py --source img.jpg \
        --task "count the mushrooms and grade harvest readiness"

Idea (https://landing.ai/blog/what-is-agentic-object-detection):
the VLM is the "agent". It looks at the photo, reasons about the question, and can call
the YOLO26 detector as a TOOL to get exact counts/boxes, then explains the result in
natural language. This is the open-vocabulary / zero-training path and complements the
trained YOLO26 detector.

Here we wire Claude (VLM) + the local YOLO26 detector as a tool. Swap in landing.ai's
VisionAgent if you prefer their hosted agent.
"""
import argparse
import base64
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def parse_args():
    p = argparse.ArgumentParser(description="Agentic VLM detection over a mushroom photo")
    p.add_argument("--source", required=True, help="image path")
    p.add_argument("--task", required=True, help="natural-language question for the agent")
    p.add_argument("--weights", default=str(ROOT / "models" / "best.pt"))
    p.add_argument("--model", default="claude-opus-4-8", help="Claude VLM model id")
    return p.parse_args()


def yolo_count(image_path, weights):
    """Tool the agent can call: returns exact mushroom count + boxes via YOLO26."""
    from ultralytics import YOLO
    w = weights if Path(weights).exists() else "yolo26n.pt"
    res = YOLO(w).predict(source=image_path, conf=0.25)[0]
    boxes = res.boxes.xyxy.cpu().numpy().round().astype(int).tolist() if res.boxes else []
    return {"count": len(boxes), "boxes_xyxy": boxes}


TOOLS = [{
    "name": "yolo_count",
    "description": "Run the trained YOLO26 detector on the image to get the exact "
                   "number of king oyster mushrooms and their bounding boxes.",
    "input_schema": {"type": "object", "properties": {}},
}]


def main():
    args = parse_args()
    try:
        import anthropic
    except ImportError:
        raise SystemExit("pip install anthropic  (and set ANTHROPIC_API_KEY)")

    client = anthropic.Anthropic()
    img_b64 = base64.standard_b64encode(Path(args.source).read_bytes()).decode()
    media = "image/png" if args.source.lower().endswith("png") else "image/jpeg"

    messages = [{
        "role": "user",
        "content": [
            {"type": "image", "source": {"type": "base64",
             "media_type": media, "data": img_b64}},
            {"type": "text", "text":
                f"You are a VisionAgent for king oyster mushrooms (杏鮑菇) growing in a "
                f"substrate bag. Task: {args.task}. Use the yolo_count tool for exact "
                f"counts, then answer in clear English + 中文."},
        ],
    }]

    # Agent loop: let the model call the detector tool, feed results back, repeat.
    while True:
        resp = client.messages.create(
            model=args.model, max_tokens=1024, tools=TOOLS, messages=messages)
        if resp.stop_reason == "tool_use":
            messages.append({"role": "assistant", "content": resp.content})
            tool_results = []
            for block in resp.content:
                if block.type == "tool_use" and block.name == "yolo_count":
                    out = yolo_count(args.source, args.weights)
                    tool_results.append({
                        "type": "tool_result", "tool_use_id": block.id,
                        "content": json.dumps(out)})
            messages.append({"role": "user", "content": tool_results})
            continue
        # Final answer.
        for block in resp.content:
            if block.type == "text":
                print(block.text)
        break


if __name__ == "__main__":
    main()
