#!/usr/bin/env python3
# =========================================================
# オフライン OCR エンジン（PP-OCRv5 / onnxocr）
#
# decompose.js からサブプロセスとして呼ばれる。
# 完全オフライン・無料（Apache-2.0）。API・課金は一切使わない。
#
# 使い方:
#   python3 ocr_engine.py --probe        → エンジンが使えるか確認（"ok" を出力）
#   python3 ocr_engine.py <image.png>    → JSON を stdout に出力
#
# 出力形式（入力画像のピクセル座標）:
#   {"lines": [{"x0":..,"y0":..,"x1":..,"y1":..,"text":"..","conf":0.97}, ...]}
# =========================================================
import sys
import json


def main() -> int:
    if len(sys.argv) < 2:
        print(json.dumps({"error": "usage: ocr_engine.py <image>|--probe"}))
        return 1

    if sys.argv[1] == "--probe":
        try:
            import cv2  # noqa: F401
            from onnxocr.onnx_paddleocr import ONNXPaddleOcr  # noqa: F401
            print("ok")
            return 0
        except Exception as e:  # pragma: no cover
            print(f"unavailable: {e}")
            return 1

    import contextlib
    import cv2
    from onnxocr.onnx_paddleocr import ONNXPaddleOcr

    img = cv2.imread(sys.argv[1])
    if img is None:
        print(json.dumps({"error": f"cannot read image: {sys.argv[1]}"}))
        return 1

    # ライブラリが stdout に出す注意書きで JSON を汚さないよう stderr へ退避する
    with contextlib.redirect_stdout(sys.stderr):
        # PP-OCRv5（既定モデル）。角度分類は水平レイアウトの図解では不要
        model = ONNXPaddleOcr(use_angle_cls=False, use_gpu=False)
        result = model.ocr(img, cls=False)

    lines = []
    for box, (text, conf) in result[0]:
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        lines.append({
            "x0": int(min(xs)), "y0": int(min(ys)),
            "x1": int(max(xs)), "y1": int(max(ys)),
            "text": text, "conf": round(float(conf), 4),
        })
    print(json.dumps({"lines": lines}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
