#!/usr/bin/env python3
# =========================================================
# オフライン OCR エンジン（PP-OCRv5 / onnxocr）
#
# decompose.js からサブプロセスとして呼ばれる。
# 完全オフライン・無料（Apache-2.0）。API・課金は一切使わない。
#
# 省メモリ設計（Render の 512MB 制約に合わせる）:
#   検出モデルは内部で長辺 960px に縮小して走るため、画像全体を
#   拡大して渡してもメモリを浪費するだけで検出結果は変わらない。
#   そこで検出は原寸のまま行い、認識(rec)に渡す行クロップだけを
#   Lanczos で 3 倍拡大する。全体拡大と同等の認識精度のまま、
#   巨大な中間画像（数百 MB）を持たなくて済む。
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

UPSCALE = 3  # 認識クロップの拡大率（小さい日本語・小書きかな対策）


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
    import copy
    import cv2
    import onnxruntime
    from onnxocr.onnx_paddleocr import ONNXPaddleOcr
    from onnxocr.predict_base import PredictBase
    from onnxocr.predict_system import sorted_boxes
    from onnxocr.utils import get_rotate_crop_image

    # onnxruntime のメモリアリーナは推論の中間テンソルを解放せず
    # 抱え込み、RSS を数百 MB 押し上げる。小さなサーバー（512MB 級）で
    # OOM になるため、アリーナを無効化した省メモリセッションに差し替える
    def _low_mem_session(self, model_dir, use_gpu):
        so = onnxruntime.SessionOptions()
        so.enable_cpu_mem_arena = False
        return onnxruntime.InferenceSession(
            model_dir, sess_options=so, providers=["CPUExecutionProvider"])
    PredictBase.get_onnx_session = _low_mem_session

    img = cv2.imread(sys.argv[1])
    if img is None:
        print(json.dumps({"error": f"cannot read image: {sys.argv[1]}"}))
        return 1

    # ライブラリが stdout に出す注意書きで JSON を汚さないよう stderr へ退避する
    with contextlib.redirect_stdout(sys.stderr):
        # PP-OCRv5（既定モデル）。角度分類は水平レイアウトの図解では不要
        model = ONNXPaddleOcr(use_angle_cls=False, use_gpu=False)
        # 検出は原寸（内部で長辺960pxに縮小される）
        dt_boxes = model.text_detector(img)
        lines = []
        if dt_boxes is not None and len(dt_boxes) > 0:
            dt_boxes = sorted_boxes(dt_boxes)
            crops = []
            for box in dt_boxes:
                crop = get_rotate_crop_image(img, copy.deepcopy(box))
                if crop is None or crop.size == 0:
                    crop = img[0:1, 0:1]
                # 認識精度のためクロップだけを拡大する
                crop = cv2.resize(crop, None, fx=UPSCALE, fy=UPSCALE,
                                  interpolation=cv2.INTER_LANCZOS4)
                crops.append(crop)
            rec_res = model.text_recognizer(crops)
            for box, (text, conf) in zip(dt_boxes, rec_res):
                if conf < 0.5:  # predict_system の drop_score と同じ既定値
                    continue
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
