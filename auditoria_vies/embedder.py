# Extração de embeddings faciais com ArcFace (buffalo_l w600k_r50) via ONNX Runtime.
import numpy as np, cv2, onnxruntime as ort

_MODEL = "models/arcface_w600k_r50.onnx"
_sess = None

def _session():
    global _sess
    if _sess is None:
        so = ort.SessionOptions()
        so.log_severity_level = 3        # silencia avisos de shape de batch
        _sess = ort.InferenceSession(_MODEL, sess_options=so,
                                     providers=["CPUExecutionProvider"])
    return _sess

def preprocess(img_rgb_uint8):
    """img_rgb_uint8: HxWx3 RGB uint8. Redimensiona p/ 112x112 e normaliza p/ ArcFace."""
    img = cv2.resize(img_rgb_uint8, (112, 112), interpolation=cv2.INTER_AREA)
    img = img.astype(np.float32)
    img = (img - 127.5) / 127.5          # [-1,1]
    img = np.transpose(img, (2, 0, 1))   # CHW
    return img[None, ...]                 # NCHW

def embed(img_rgb_uint8):
    s = _session()
    out = s.run(None, {s.get_inputs()[0].name: preprocess(img_rgb_uint8)})[0][0]
    n = np.linalg.norm(out)
    return out / (n + 1e-9)

def embed_batch(imgs_rgb_uint8):
    s = _session()
    batch = np.concatenate([preprocess(im) for im in imgs_rgb_uint8], axis=0)
    out = s.run(None, {s.get_inputs()[0].name: batch})[0]
    out = out / (np.linalg.norm(out, axis=1, keepdims=True) + 1e-9)
    return out

def cosine(a, b):
    return float(np.dot(a, b))
