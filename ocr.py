import numpy as np
from typing import List, Tuple

_ocr_instance = None


def get_ocr(use_gpu: bool = False):
    """
    Initialize PaddleOCR, compatible with both v2.x and v3.x APIs.
    NOTE: GPU mode requires a paddlepaddle-gpu build matching your CUDA version.
    CUDA 13.x is not yet supported by paddlepaddle-gpu; use CPU mode instead.
    """
    global _ocr_instance
    if _ocr_instance is None:
        # Must disable oneDNN BEFORE importing PaddlePaddle.
        # Environment variables are ignored by PaddlePaddle 3.x at runtime;
        # paddle.set_flags() is the only reliable programmatic approach.
        try:
            import paddle
            paddle.set_flags({
                "FLAGS_use_mkldnn": False,
                "FLAGS_enable_pir_api": False,
            })
        except Exception:
            pass

        # paddlex 的 static_infer.py 在创建推理器时会显式调用
        # config.enable_new_ir(True)，覆盖全局 FLAGS。新 IR（PIR）在
        # Windows CPU 模式下与 oneDNN 存在兼容 bug，触发
        # ConvertPirAttribute2RuntimeAttribute 报错。
        # 解决方法：把 PP-OCRv5 检测模型加入 NEWIR_BLOCKLIST，
        # 让 paddlex 对这些模型调用 config.enable_new_ir(False)。
        try:
            from paddlex.inference.utils.new_ir_blocklist import NEWIR_BLOCKLIST
            for _det_model in (
                "PP-OCRv5_server_det",
                "PP-OCRv5_mobile_det",
                "PP-OCRv4_server_det",
                "PP-OCRv4_mobile_det",
            ):
                if _det_model not in NEWIR_BLOCKLIST:
                    NEWIR_BLOCKLIST.append(_det_model)
        except Exception:
            pass

        from paddleocr import PaddleOCR

        device_str = "gpu" if use_gpu else "cpu"

        # PaddleOCR 3.x: disable the heavy document-preprocessing sub-models
        # (orientation classifier, unwarping, textline orientation) — these are
        # unnecessary for clean game UI screenshots and some of them trigger the
        # oneDNN PIR bug in PP-OCRv5_server_det.
        v3_kwargs = dict(
            lang="ch",
            device=device_str,
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
        )
        try:
            _ocr_instance = PaddleOCR(**v3_kwargs)
        except (TypeError, ValueError):
            # Some params may not exist in this v3 sub-version; drop extras
            try:
                _ocr_instance = PaddleOCR(lang="ch", device=device_str)
            except (TypeError, ValueError):
                # Fall back to PaddleOCR 2.x API
                _ocr_instance = PaddleOCR(
                    use_angle_cls=False,
                    lang="ch",
                    use_gpu=use_gpu,
                    show_log=False,
                )

        print(f"  OCR 后端：{device_str.upper()}")

    return _ocr_instance


def _parse_result(result) -> List[Tuple[str, float]]:
    """
    Parse raw PaddleOCR output into (text, confidence) pairs.
    Handles both v2 and v3 result formats.
    """
    lines = []

    if not result:
        return lines

    # ── Debug: print raw result type so we can diagnose format mismatches ──
    page = result[0]
    print(f"  [DEBUG] result type={type(result).__name__}, page type={type(page).__name__}")

    # ── PaddleOCR v3 (PaddleX pipeline): page is an OCRResult / dict-like ──
    # Try attribute-based access first (rec_texts / rec_scores)
    rec_texts = None
    rec_scores = None
    try:
        rec_texts = page["rec_texts"] if isinstance(page, dict) else getattr(page, "rec_texts", None)
        rec_scores = page["rec_scores"] if isinstance(page, dict) else getattr(page, "rec_scores", None)
    except Exception:
        pass

    if rec_texts is not None and rec_scores is not None:
        for text, conf in zip(rec_texts, rec_scores):
            try:
                lines.append((str(text).strip(), float(conf)))
            except (TypeError, ValueError):
                continue
        return lines

    # ── PaddleOCR v2 / v3 compat: page is a list of [bbox, [text, conf]] ──
    if page is None:
        return lines

    try:
        items = list(page)
    except TypeError:
        return lines

    for item in items:
        try:
            # Standard format: [bbox, [text, confidence]]
            text = item[1][0]
            conf = float(item[1][1])
            lines.append((text.strip(), conf))
        except (IndexError, TypeError, ValueError):
            continue

    return lines


def recognize(
    img: np.ndarray,
    use_gpu: bool = True,
    confidence_threshold: float = 0.7,
) -> List[Tuple[str, float]]:
    """
    Run OCR on the given image.
    Returns list of (text, confidence) tuples where confidence >= threshold.
    """
    ocr = get_ocr(use_gpu)
    result = ocr.ocr(img)

    return [
        (text, conf)
        for text, conf in _parse_result(result)
        if conf >= confidence_threshold
    ]


def _parse_result_with_pos(result) -> List[Tuple[str, float, int, int]]:
    """
    Parse raw PaddleOCR output into (text, confidence, center_x, center_y) tuples.
    Handles both v3 attribute format (rec_texts/dt_polys) and v2 list format.
    """
    lines: List[Tuple[str, float, int, int]] = []
    if not result:
        return lines

    page = result[0]

    # PaddleOCR v3: try rec_texts / rec_scores / dt_polys
    rec_texts = None
    rec_scores = None
    dt_polys = None
    try:
        rec_texts = page["rec_texts"] if isinstance(page, dict) else getattr(page, "rec_texts", None)
        rec_scores = page["rec_scores"] if isinstance(page, dict) else getattr(page, "rec_scores", None)
        dt_polys  = page["dt_polys"]  if isinstance(page, dict) else getattr(page, "dt_polys",  None)
    except Exception:
        pass

    if rec_texts is not None and rec_scores is not None and dt_polys is not None:
        for text, conf, poly in zip(rec_texts, rec_scores, dt_polys):
            try:
                pts = np.array(poly, dtype=float)
                cx = int(pts[:, 0].mean())
                cy = int(pts[:, 1].mean())
                lines.append((str(text).strip(), float(conf), cx, cy))
            except Exception:
                continue
        return lines

    # PaddleOCR v2 / v3 compat: list of [bbox, [text, conf]]
    if page is None:
        return lines
    try:
        items = list(page)
    except TypeError:
        return lines

    for item in items:
        try:
            bbox = item[0]
            text = item[1][0]
            conf = float(item[1][1])
            pts = np.array(bbox, dtype=float)
            cx = int(pts[:, 0].mean())
            cy = int(pts[:, 1].mean())
            lines.append((text.strip(), conf, cx, cy))
        except (IndexError, TypeError, ValueError):
            continue

    return lines


def recognize_with_positions(
    img: np.ndarray,
    use_gpu: bool = False,
    confidence_threshold: float = 0.7,
) -> List[Tuple[str, float, int, int]]:
    """
    Run OCR and return (text, confidence, center_x, center_y) for each result.
    Coordinates are in pixels relative to the input image.
    """
    ocr = get_ocr(use_gpu)
    result = ocr.ocr(img)
    return [
        (text, conf, cx, cy)
        for text, conf, cx, cy in _parse_result_with_pos(result)
        if conf >= confidence_threshold
    ]
