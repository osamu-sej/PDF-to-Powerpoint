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
#   Lanczos で拡大する。
#
# 誤読対策:
#   - 検出枠は行末の文字を欠くことがあるため、行高に応じて左右に
#     広げてから認識する（隣の行があればその手前まで）
#   - 同じクロップを 3 倍と 1.5 倍の 2 スケールで認識し、両方の
#     読みを返す。呼び出し側は「2 つの読みが一致するか」を
#     誤読検出のシグナルとして使う（劣化画像の誤読はスケールに
#     敏感で、正しい読みはスケールに安定している）
#
# 使い方:
#   python3 ocr_engine.py --probe        → エンジンが使えるか確認（"ok" を出力）
#   python3 ocr_engine.py <image.png>    → JSON を stdout に出力
#
# 出力形式（入力画像のピクセル座標）:
#   {"lines":[{x0,y0,x1,y1,text,conf,text2,conf2}, ...]}
#   text/conf は 3 倍スケールの読み、text2/conf2 は 1.5 倍の読み
# =========================================================
import sys
import json

UPSCALE_MAIN = 3.0   # 主認識スケール
UPSCALE_ALT = 1.5    # 照合用の副認識スケール
EXPAND_X = 0.35      # 検出枠の左右拡張（行高比）
EXPAND_Y = 0.10      # 検出枠の上下拡張（行高比）


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
    import onnxruntime
    from onnxocr.onnx_paddleocr import ONNXPaddleOcr
    from onnxocr.predict_base import PredictBase
    from onnxocr.predict_system import sorted_boxes

    # onnxruntime のメモリアリーナは中間テンソルを抱え込み RSS を
    # 数百 MB 押し上げるため無効化する（512MB 級サーバー対策）
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
    H, W = img.shape[:2]

    with contextlib.redirect_stdout(sys.stderr):
        model = ONNXPaddleOcr(use_angle_cls=False, use_gpu=False)
        dt_boxes = model.text_detector(img)

    lines = []
    if dt_boxes is not None and len(dt_boxes) > 0:
        dt_boxes = sorted_boxes(dt_boxes)
        # 軸平行の矩形に落とす（水平レイアウトの図解が対象）
        rects = []
        for box in dt_boxes:
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            rects.append([int(min(xs)), int(min(ys)), int(max(xs)), int(max(ys))])

        # 左右拡張: 同じ行帯にある隣の枠の手前まで
        def y_overlap(a, b):
            iv = min(a[3], b[3]) - max(a[1], b[1])
            return iv / max(1, min(a[3] - a[1], b[3] - b[1]))

        expanded = []
        for i, r in enumerate(rects):
            h = max(1, r[3] - r[1])
            ex = int(h * EXPAND_X)
            ey = int(h * EXPAND_Y)
            left_lim, right_lim = 0, W
            for j, q in enumerate(rects):
                if i == j or y_overlap(r, q) < 0.5:
                    continue
                if q[2] <= r[0]:
                    left_lim = max(left_lim, q[2] + 2)
                if q[0] >= r[2]:
                    right_lim = min(right_lim, q[0] - 2)
            x0 = max(left_lim, r[0] - ex)
            x1 = min(right_lim, r[2] + ex)
            if x1 <= x0:
                x0, x1 = r[0], r[2]
            y0 = max(0, r[1] - ey)
            y1 = min(H, r[3] + ey)
            expanded.append([x0, y0, x1, y1])

        crops = []
        for (x0, y0, x1, y1) in expanded:
            crop = img[y0:y1, x0:x1]
            if crop.size == 0:
                crop = img[0:1, 0:1]
            crops.append(crop)

        def rec_at(scale):
            ups = [cv2.resize(c, None, fx=scale, fy=scale,
                              interpolation=cv2.INTER_LANCZOS4) for c in crops]
            with contextlib.redirect_stdout(sys.stderr):
                return model.text_recognizer(ups)

        r_main = rec_at(UPSCALE_MAIN)
        r_alt = rec_at(UPSCALE_ALT)

        for (x0, y0, x1, y1), (t3, c3), (t2, c2) in zip(expanded, r_main, r_alt):
            if not t3 or c3 < 0.5:
                continue
            lines.append({
                "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                "text": t3, "conf": round(float(c3), 4),
                "text2": t2, "conf2": round(float(c2), 4),
            })
    print(json.dumps({"lines": lines}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
