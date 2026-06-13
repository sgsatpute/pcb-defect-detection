"""
PCB Defect Detection - Streamlit Demo App

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Usage:
    1. Keep the trained best.pt in this folder, or upload a .pt file.
    2. Upload any PCB image.
    3. Adjust confidence and IOU thresholds as needed.
"""

from __future__ import annotations

import io
import tempfile
from collections import defaultdict
from pathlib import Path

import cv2
import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image


st.set_page_config(
    page_title="PCB Defect Detector",
    layout="wide",
    initial_sidebar_state="expanded",
)


FALLBACK_CLASS_NAMES = ["open", "short", "mousebite", "spur", "copper", "pin-hole"]

CLASS_COLORS = {
    "open": "#e74c3c",
    "short": "#3498db",
    "mousebite": "#2ecc71",
    "spur": "#f39c12",
    "copper": "#9b59b6",
    "pin-hole": "#1abc9c",
}

DEFAULT_MODEL_PATHS = (
    Path("best.pt"),
    Path("models/best.pt"),
    Path("results/pcb_results/weights/best.pt"),
)


st.markdown(
    """
<style>
    [data-testid="stSidebar"] {background: #0d1117;}
    [data-testid="stSidebar"] * {color: #e6edf3 !important;}
    .stApp {background: #161b22;}
    .block-container {padding-top: 1.5rem;}
    h1 {color: #e6edf3 !important;}
    h2, h3 {color: #c9d1d9 !important;}
    p, li {color: #8b949e;}
    .stSlider label {color: #c9d1d9 !important;}
    .metric-card {
        background: #21262d;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 14px 18px;
        margin: 6px 0;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .metric-dot {
        width: 14px;
        height: 14px;
        border-radius: 50%;
        flex-shrink: 0;
    }
    .metric-label {
        color: #8b949e;
        font-size: 13px;
        text-transform: uppercase;
        letter-spacing: .5px;
    }
    .metric-value {
        color: #e6edf3;
        font-size: 22px;
        font-weight: 700;
    }
    .section-title {
        color: #58a6ff;
        font-size: 15px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: .8px;
        margin: 18px 0 8px;
        padding-bottom: 6px;
        border-bottom: 1px solid #30363d;
    }
    .badge {
        display: inline-block;
        padding: 2px 9px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        color: white;
        margin: 2px 3px;
    }
</style>
""",
    unsafe_allow_html=True,
)


@st.cache_resource(show_spinner=False)
def load_model(model_path: str):
    from ultralytics import YOLO

    return YOLO(model_path)


def default_model_path() -> str:
    for path in DEFAULT_MODEL_PATHS:
        if path.exists():
            return str(path)
    return "best.pt"


def get_class_name(model, class_id: int) -> str:
    names = getattr(model, "names", None)
    if isinstance(names, dict) and class_id in names:
        return str(names[class_id])
    if isinstance(names, list) and class_id < len(names):
        return str(names[class_id])
    if class_id < len(FALLBACK_CLASS_NAMES):
        return FALLBACK_CLASS_NAMES[class_id]
    return f"class_{class_id}"


def class_names_for_legend(model) -> list[str]:
    names = getattr(model, "names", None) if model is not None else None
    if isinstance(names, dict):
        return [str(names[key]) for key in sorted(names)]
    if isinstance(names, list):
        return [str(name) for name in names]
    return FALLBACK_CLASS_NAMES


def render_metric_card(label: str, value: int) -> None:
    color = CLASS_COLORS.get(label, "#666")
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-dot" style="background:{color};"></div>
            <div>
                <div class="metric-label">{label}</div>
                <div class="metric-value">{value}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


with st.sidebar:
    st.markdown("## PCB Defect Detector")
    st.markdown("---")

    st.markdown("### Model")
    model_source = st.radio(
        "Load model from",
        ["Local file", "Upload .pt file"],
        label_visibility="collapsed",
    )

    model = None
    if model_source == "Local file":
        local_path = Path(st.text_input("Model path", value=default_model_path()))
        if local_path.exists():
            try:
                with st.spinner("Loading model..."):
                    model = load_model(str(local_path))
                st.success("Model loaded")
            except Exception as exc:
                st.error(f"Model could not be loaded: {exc}")
        else:
            st.warning("Model file not found. Keep best.pt beside app.py or upload a .pt file.")
    else:
        uploaded_model = st.file_uploader("Upload YOLO model", type=["pt"])
        if uploaded_model:
            with tempfile.NamedTemporaryFile(suffix=".pt", delete=False) as tmp:
                tmp.write(uploaded_model.read())
                tmp_path = tmp.name
            try:
                with st.spinner("Loading model..."):
                    model = load_model(tmp_path)
                st.success("Model loaded")
            except Exception as exc:
                st.error(f"Uploaded model could not be loaded: {exc}")

    st.markdown("---")
    st.markdown("### Inference Settings")
    conf_thresh = st.slider(
        "Confidence threshold",
        min_value=0.05,
        max_value=0.95,
        value=0.25,
        step=0.05,
        help="Minimum confidence for a detection to be shown.",
    )
    iou_thresh = st.slider(
        "IOU threshold (NMS)",
        min_value=0.10,
        max_value=0.95,
        value=0.45,
        step=0.05,
        help="Higher values keep more overlapping boxes before suppression.",
    )

    st.markdown("---")
    st.markdown("### Defect Classes")
    for cls in class_names_for_legend(model):
        color = CLASS_COLORS.get(cls, "#666")
        st.markdown(
            f'<span class="badge" style="background:{color};">{cls}</span>',
            unsafe_allow_html=True,
        )

    st.markdown("---")
    st.caption("YOLOv8s | DeepPCB Dataset | Streamlit demo")


st.title("PCB Defect Detection")
st.markdown(
    "Upload a Printed Circuit Board image to detect manufacturing defects using "
    "a YOLOv8s model trained on the DeepPCB dataset."
)

uploaded_img = st.file_uploader(
    "Upload a PCB image",
    type=["jpg", "jpeg", "png", "bmp"],
    help="Drag and drop or click to browse.",
)

if uploaded_img is None:
    st.markdown("---")
    col_l, col_r = st.columns(2)
    with col_l:
        st.info(
            "**Getting started**\n\n"
            "1. Load `best.pt` in the sidebar.\n"
            "2. Upload a PCB image.\n"
            "3. Adjust thresholds to filter detections.\n\n"
            "Supported defect types: "
            + ", ".join(f"`{c}`" for c in FALLBACK_CLASS_NAMES)
        )
    with col_r:
        st.markdown(
            """
            **Model summary**

            | Item | Detail |
            |---|---|
            | Architecture | YOLOv8s |
            | Dataset | DeepPCB |
            | Classes | 6 defect types |
            | Input size | 640 x 640 |
            | Best mAP@0.5 | 0.99376 |
            """
        )
    st.stop()

if model is None:
    st.error("Load a model first using the sidebar.")
    st.stop()

file_bytes = np.frombuffer(uploaded_img.read(), np.uint8)
img_bgr = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
if img_bgr is None:
    st.error("The uploaded file could not be decoded as an image.")
    st.stop()

img_rgb = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)

with st.spinner("Running inference..."):
    result = model.predict(img_bgr, conf=conf_thresh, iou=iou_thresh, verbose=False)[0]
    annotated_bgr = result.plot(line_width=2)
    annotated_rgb = cv2.cvtColor(annotated_bgr, cv2.COLOR_BGR2RGB)

st.markdown("---")
col_orig, col_det = st.columns(2)

with col_orig:
    st.markdown('<p class="section-title">Original Image</p>', unsafe_allow_html=True)
    st.image(img_rgb, use_container_width=True)
    height, width = img_bgr.shape[:2]
    st.caption(f"{width} x {height} px")

with col_det:
    st.markdown('<p class="section-title">Detection Result</p>', unsafe_allow_html=True)
    st.image(annotated_rgb, use_container_width=True)
    total = len(result.boxes)
    st.caption(
        f"{total} defect(s) detected | conf >= {conf_thresh:.2f} | IOU <= {iou_thresh:.2f}"
    )

st.markdown("---")
st.markdown('<p class="section-title">Detection Breakdown</p>', unsafe_allow_html=True)

boxes = result.boxes
if len(boxes) == 0:
    st.info("No defects were detected above the current confidence threshold.")
else:
    counts: dict[str, int] = defaultdict(int)
    max_conf: dict[str, float] = defaultdict(float)
    all_conf: dict[str, list[float]] = defaultdict(list)

    for box in boxes:
        class_id = int(box.cls.item())
        conf = float(box.conf.item())
        class_name = get_class_name(model, class_id)
        counts[class_name] += 1
        max_conf[class_name] = max(max_conf[class_name], conf)
        all_conf[class_name].append(conf)

    metric_columns = st.columns(min(len(counts), 6))
    for index, (class_name, count) in enumerate(sorted(counts.items())):
        with metric_columns[index % len(metric_columns)]:
            render_metric_card(class_name, count)

    rows = []
    for class_name in sorted(counts):
        confidences = all_conf[class_name]
        rows.append(
            {
                "Defect Type": class_name,
                "Count": counts[class_name],
                "Max Confidence": f"{max_conf[class_name]:.3f}",
                "Avg Confidence": f"{sum(confidences) / len(confidences):.3f}",
            }
        )
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

st.markdown("---")
pil_annotated = Image.fromarray(annotated_rgb)
buffer = io.BytesIO()
pil_annotated.save(buffer, format="PNG")
st.download_button(
    label="Download annotated image",
    data=buffer.getvalue(),
    file_name=f"pcb_detection_{Path(uploaded_img.name).stem}.png",
    mime="image/png",
)
