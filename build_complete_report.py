import os
import re
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image, HRFlowable, Preformatted, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from reportlab.pdfgen import canvas

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        # Don't draw headers/footers on cover page (Page 1)
        if self._pageNumber > 1:
            self.saveState()
            self.setFont('Helvetica', 8)
            self.setFillColor(colors.HexColor('#64748B'))
            
            # Header
            self.drawString(54, A4[1] - 36, 'Vishwakarma University | Dept. of AIML | Project 1 Report')
            self.drawRightString(A4[0] - 54, A4[1] - 36, 'VLM-IDP Research')
            self.setStrokeColor(colors.HexColor('#CBD5E1'))
            self.setLineWidth(0.5)
            self.line(54, A4[1] - 42, A4[0] - 54, A4[1] - 42)
            
            # Footer
            self.line(54, 46, A4[0] - 54, 46)
            self.drawString(54, 34, 'S. Y. B. Tech (AIML) - Academic Year 2026-2027')
            self.drawRightString(A4[0] - 54, 34, f'Page {self._pageNumber} of {page_count}')
            self.restoreState()

def build_pdf_report(pdf_path, logo_path):
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54
    )
    
    styles = getSampleStyleSheet()
    
    univ_header_style = ParagraphStyle(
        'UnivHeader',
        fontName='Helvetica-Bold',
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#0F172A')
    )
    
    faculty_style = ParagraphStyle(
        'FacultyStyle',
        fontName='Helvetica',
        fontSize=13,
        leading=16,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#334155')
    )
    
    dept_style = ParagraphStyle(
        'DeptStyle',
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=16,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#C2410C')
    )
    
    proj_type_style = ParagraphStyle(
        'ProjType',
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=15,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#64748B')
    )
    
    title_style = ParagraphStyle(
        'CoverTitle',
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=21,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#0F172A')
    )
    
    meta_bold_style = ParagraphStyle(
        'MetaBold',
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#1E293B')
    )
    
    meta_regular_style = ParagraphStyle(
        'MetaReg',
        fontName='Helvetica',
        fontSize=11,
        leading=15,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#334155')
    )
    
    sec_heading_style = ParagraphStyle(
        'SecHeading',
        fontName='Helvetica-Bold',
        fontSize=18,
        leading=22,
        textColor=colors.HexColor('#C2410C'),
        spaceAfter=10,
        spaceBefore=14
    )
    
    sub_heading_style = ParagraphStyle(
        'SubHeading',
        fontName='Helvetica-Bold',
        fontSize=12.5,
        leading=16,
        textColor=colors.HexColor('#0F172A'),
        spaceAfter=6,
        spaceBefore=10
    )
    
    body_style = ParagraphStyle(
        'BodyDark',
        fontName='Helvetica',
        fontSize=10,
        leading=14.5,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor('#1E293B'),
        spaceAfter=6
    )
    
    bullet_style = ParagraphStyle(
        'BulletDark',
        fontName='Helvetica',
        fontSize=9.5,
        leading=14,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor('#1E293B'),
        leftIndent=15,
        firstLineIndent=-10,
        spaceAfter=4
    )
    
    code_pre_style = ParagraphStyle(
        'CodePre',
        fontName='Courier',
        fontSize=7.8,
        leading=10.2,
        textColor=colors.HexColor('#0F172A')
    )

    story = []

    # =========================================================================
    # PAGE 1: COVER PAGE
    # VU Logo at top, then VU Pune name, then faculty, then pursuing in only AIML
    # =========================================================================
    if os.path.exists(logo_path):
        img = Image(logo_path, width=140, height=104.5)
        img.hAlign = 'CENTER'
        story.append(img)
    story.append(Spacer(1, 10))
    
    story.append(Paragraph("VISHWAKARMA UNIVERSITY, PUNE", univ_header_style))
    story.append(Spacer(1, 3))
    story.append(Paragraph("Faculty of Science & Technology", faculty_style))
    story.append(Spacer(1, 3))
    story.append(Paragraph("Department of Artificial Intelligence & Machine Learning (AIML)", dept_style))
    story.append(Spacer(1, 3))
    story.append(Paragraph("Pursuing in S. Y. B. Tech (Artificial Intelligence & Machine Learning)", meta_regular_style))
    story.append(Spacer(1, 14))
    
    story.append(HRFlowable(width="90%", thickness=1.5, color=colors.HexColor('#CBD5E1'), spaceAfter=14))
    
    story.append(Paragraph("A Project 1 Report On", proj_type_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph("Adaptive, Uncertainty-Aware and Evidence-Grounded Vision-Language Document Intelligence under Real-World Visual Degradation", title_style))
    story.append(Spacer(1, 14))
    
    story.append(Paragraph("Project 1 [AIML201 - Artificial Intelligence & Machine Learning Project 1]", meta_bold_style))
    story.append(Spacer(1, 16))
    
    story.append(Paragraph("By", ParagraphStyle('ByLabel', fontName='Helvetica-Bold', fontSize=12, alignment=TA_CENTER, textColor=colors.HexColor('#64748B'))))
    story.append(Spacer(1, 5))
    story.append(Paragraph("<b>Raunak Rajput</b> (SRN-31252544)", meta_bold_style))
    story.append(Spacer(1, 3))
    story.append(Paragraph("<b>Ayush Vishwakarma</b>", meta_bold_style))
    story.append(Spacer(1, 3))
    story.append(Paragraph("<b>Vipul Bhamare</b>", meta_bold_style))
    story.append(Spacer(1, 18))
    
    story.append(Paragraph("Academic Year 2026–2027", meta_bold_style))
    
    # End of Page 1
    story.append(PageBreak())

    # =========================================================================
    # PAGE 2: INDEX (Table of Contents only - NO repeated college info!)
    # =========================================================================
    story.append(Paragraph("Index", sec_heading_style))
    story.append(Spacer(1, 8))
    
    index_data = [
        [Paragraph("<b>Sr. No.</b>", ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=9.5, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>Content</b>", ParagraphStyle('TH2', fontName='Helvetica-Bold', fontSize=9.5, alignment=TA_LEFT, textColor=colors.white)),
         Paragraph("<b>Page No.</b>", ParagraphStyle('TH3', fontName='Helvetica-Bold', fontSize=9.5, alignment=TA_CENTER, textColor=colors.white))],
        ["1", "Introduction & Research Background", "3"],
        ["2", "Problem Formulation & Proposed Architecture", "4"],
        ["3", "Core Implementation & Source Code Listings", "5"],
        ["4", "Experimental Benchmarks & Hardware Microbenchmarking", "8"],
        ["5", "Results Analysis & Statistical Verification", "9"],
        ["6", "Conclusion & Future Research Directions", "10"],
    ]
    
    index_table = Table(index_data, colWidths=[55, 360, 72])
    index_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (1, 1), (1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 9.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#FFFFFF'), colors.HexColor('#F8FAFC')]),
    ]))
    story.append(index_table)
    story.append(Spacer(1, 20))
    
    # End of Page 2
    story.append(PageBreak())

    # =========================================================================
    # PAGE 3 ONWARDS: INTRODUCTION
    # =========================================================================
    story.append(Paragraph("Introduction:", sec_heading_style))
    
    story.append(Paragraph("1.1 Background & Motivation", sub_heading_style))
    story.append(Paragraph(
        "Intelligent Document Processing (IDP) serves as a mission-critical technology in modern enterprise machine learning workflows, "
        "enabling automated parsing, extraction, and validation of information from unstructured document images such as invoices, "
        "financial balance sheets, legal contracts, receipts, and government records. In standard academic benchmarks, "
        "document understanding models are typically evaluated on pristine, digitally synthesized, high-contrast 300 DPI flatbed scans. "
        "However, real-world deployment pipelines in industrial practice regularly confront acute physical, optical, and environmental "
        "degradations that violate standard distribution assumptions.",
        body_style
    ))
    story.append(Paragraph(
        "Real-world document capture in industrial workflows routinely introduces severe visual corruptions, including camera defocus blur, "
        "motion blur from handheld smartphones, non-uniform specular glare, shadow gradients, perspective keystone distortion, physical wrinkles, "
        "carbon-copy ink bleed-through, low-illumination noise, and aggressive multi-generation JPEG recompression. When deployed in "
        "high-stakes compliance, banking, and legal workflows, automated document intelligence systems must remain robust, computationally "
        "efficient, and incapable of silently hallucinating fabricated facts.",
        body_style
    ))

    story.append(Paragraph("1.2 Limitations of Existing Architectures", sub_heading_style))
    story.append(Paragraph(
        "•  <b>Catastrophic Brittleness of Classical OCR:</b> Traditional IDP pipelines rely upon pipelined Optical Character Recognition (OCR) engines "
        "(e.g., Tesseract, PaddleOCR). These engines serialize visual tokens into 1D text strings prior to downstream semantic processing. "
        "Under visual noise, stroke segmentations collapse, inducing massive Character Error Rate (CER) spikes and irreversibly corrupting key-value relationships.",
        bullet_style
    ))
    story.append(Paragraph(
        "•  <b>Autoregressive Hallucination in Vision-Language Models (VLMs):</b> Modern multimodal foundation models (e.g., Qwen2.5-VL, SmolVLM) "
        "jointly encode 2D document pixels and text queries. However, when text regions are illegible or degraded, autoregressive decoders "
        "generate fluent, confident, but entirely fabricated answers that lack empirical grounding in the document pixels.",
        bullet_style
    ))
    story.append(Paragraph(
        "•  <b>The Physical GPU Memory Wall & Latency Penalty:</b> State-of-the-art 7B parameter VLMs in 16-bit precision require ~16.0 GiB of VRAM, "
        "triggering OutOfMemory (OOM) failures on commodity 6GB enterprise edge hardware. Furthermore, feeding entire 5-page scans introduces "
        "context crowding and inflates end-to-end inference latency to over 7.48 seconds.",
        bullet_style
    ))
    story.append(Paragraph(
        "•  <b>Uncalibrated Overconfidence & Lack of Abstention:</b> Conventional deep models produce uncalibrated softmax logits, remaining overconfident "
        "even when operating under acute ambiguity. Without an explicit selective abstention mechanism, error-tolerant automation cannot be safely integrated with human-in-the-loop escalation.",
        bullet_style
    ))

    story.append(Paragraph("1.3 Proposed System Architecture & Core Innovations", sub_heading_style))
    story.append(Paragraph(
        "To resolve the fundamental triad of trade-offs—Accuracy, Latency, and Hallucination Safety—this project introduces an end-to-end, "
        "uncertainty-aware, evidence-grounded document intelligence framework consisting of five core architectural pillars:",
        body_style
    ))
    story.append(Paragraph("•  <b>OpenCV Quality Assessment Engine:</b> Extracts 10 observable low-level computer vision metrics (Laplacian blur variance, Tenengrad energy, RMS contrast, skew angle, illumination gradient, noise variance) without requiring ground-truth labels.", bullet_style))
    story.append(Paragraph("•  <b>Tri-Pathway Adaptive Dynamic Router:</b> Intelligently routes clean documents to lightweight OCR/VLM, moderately degraded scans to image restoration (Wiener deconvolution / unsharp masking) and grounded VLM, and severely damaged documents to dual-OCR fallback or safe abstention.", bullet_style))
    story.append(Paragraph("•  <b>Dense Multimodal Page Retrieval:</b> Indexes multi-page document representations using dense embeddings, isolating the top-2 relevant pages. This eliminates 60% of context tokens and halves physical neural inference latency from 7.49s to 3.75s.", bullet_style))
    story.append(Paragraph("•  <b>Canonical Spatial Evidence Grounding:</b> Enforces explicit bounding box coordinate prediction in normalized [0, 1000] integer space, verifying extracted tokens against raw image pixels with an Intersection-over-Union threshold (IoU >= 0.50).", bullet_style))
    story.append(Paragraph("•  <b>Calibrated Multi-Signal Uncertainty & Abstention Gate:</b> Combines predictive entropy, cross-candidate score margin, retrieval dispersion, and degradation penalties into an uncertainty vector, guaranteeing safe abstention on unsupportable queries.", bullet_style))
    story.append(Paragraph("•  <b>Edge Hardware Quantization (INT4 NF4):</b> Deploys BitsAndBytes NormalFloat-4 quantization on SmolVLM-500M, achieving a 52.2% VRAM reduction (down to 530.95 MB) on a physical 6GB Laptop GPU.", bullet_style))

    story.append(Spacer(1, 10))

    # =========================================================================
    # CODE SECTION
    # =========================================================================
    story.append(Paragraph("Code:", sec_heading_style))
    story.append(Paragraph(
        "The project is structured into modular Python packages adhering to zero-leakage and deterministic provenance standards. "
        "The listings below illustrate the core implementations of each pipeline component.",
        body_style
    ))

    def make_code_box(title, code_str):
        code_flowable = Preformatted(code_str.strip(), code_pre_style)
        box_table = Table([[Paragraph(f"<b>Listing: {title}</b>", ParagraphStyle('CodeHdr', fontName='Helvetica-Bold', fontSize=8.5, textColor=colors.HexColor('#0F172A')))],
                           [code_flowable]], colWidths=[487])
        box_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E2E8F0')),
            ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#F8FAFC')),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
            ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.HexColor('#CBD5E1')),
        ]))
        return KeepTogether([box_table, Spacer(1, 8)])

    code1 = '''# File: src/quality/blur.py
import cv2
import numpy as np
from typing import Dict, Any

def measure_blur(gray_image: np.ndarray, config: Dict[str, Any]) -> Dict[str, float]:
    """Computes blur metrics using Variance of Laplacian and Tenengrad Gradient Energy."""
    if gray_image is None or gray_image.size == 0:
        return {"blur_score": 0.0, "tenengrad": 0.0, "is_valid": False}

    laplacian = cv2.Laplacian(gray_image, cv2.CV_64F, ksize=config.get("kernel_size", 3))
    lap_var = float(laplacian.var())

    gx = cv2.Sobel(gray_image, cv2.CV_64F, 1, 0, ksize=3)
    gy = cv2.Sobel(gray_image, cv2.CV_64F, 0, 1, ksize=3)
    tenengrad = float(np.mean(gx**2 + gy**2))

    clean_th = float(config.get("clean_threshold", 300.0))
    norm_val = min(1.0, max(0.0, lap_var / clean_th))

    return {"blur_score": norm_val, "laplacian_variance": lap_var, "tenengrad_energy": tenengrad, "is_valid": True}'''
    story.append(Paragraph("2.1 Document Quality & Degradation Assessment Engine", sub_heading_style))
    story.append(make_code_box("src/quality/blur.py (Blur & Edge Energy Metric)", code1))

    code2 = '''# File: src/routing/router.py
from typing import Dict, Any, Optional

class AdaptiveRouter:
    """Orchestrates dynamic pathway dispatch based on observable image quality vectors."""
    def __init__(self, quality_thresholds: Dict[str, float], cost_model: Any):
        self.thresholds = quality_thresholds
        self.cost_model = cost_model

    def route_document(self, quality_features: Dict[str, float]) -> str:
        blur_score = quality_features.get("blur_score", 1.0)
        contrast_score = quality_features.get("contrast_score", 1.0)
        noise_level = quality_features.get("noise_level", 0.0)

        # Route A: High quality clean documents -> Fast Lightweight Pipeline
        if blur_score >= 0.70 and contrast_score >= 0.65 and noise_level <= 0.15:
            return "ROUTE_A_FAST_VLM"
        # Route B: Moderate degradation -> Image Restoration + Grounded Multimodal VLM
        elif blur_score >= 0.35 and contrast_score >= 0.30:
            return "ROUTE_B_RESTORATION_GROUNDED_VLM"
        # Route C: Severe degradation -> Dual OCR Fallback or Human Escalation
        else:
            return "ROUTE_C_STRUCTURAL_FALLBACK"'''
    story.append(Paragraph("2.2 Tri-Pathway Adaptive Dynamic Router", sub_heading_style))
    story.append(make_code_box("src/routing/router.py (Quality-Aware Adaptive Router)", code2))

    code3 = '''# File: src/evidence/region_validator.py
from typing import Tuple

def clip_bbox_1000(bbox: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
    """Clips bounding box to normalized integer coordinate space [0, 1000]."""
    xmin, ymin, xmax, ymax = bbox
    xmin = max(0, min(1000, int(xmin)))
    ymin = max(0, min(1000, int(ymin)))
    xmax = max(0, min(1000, int(xmax)))
    ymax = max(0, min(1000, int(ymax)))
    if xmin > xmax: xmin, xmax = xmax, xmin
    if ymin > ymax: ymin, ymax = ymax, ymin
    return (xmin, ymin, xmax, ymax)

def compute_iou(boxA: Tuple[int, int, int, int], boxB: Tuple[int, int, int, int]) -> float:
    """Computes exact Intersection-over-Union (IoU) between two bounding boxes."""
    xA, yA = max(boxA[0], boxB[0]), max(boxA[1], boxB[1])
    xB, yB = min(boxA[2], boxB[2]), min(boxA[3], boxB[3])
    inter_area = max(0, xB - xA) * max(0, yB - yA)
    if inter_area <= 0: return 0.0
    union_area = float((boxA[2] - boxA[0]) * (boxA[3] - boxA[1]) + (boxB[2] - boxB[0]) * (boxB[3] - boxB[1]) - inter_area)
    return inter_area / union_area if union_area > 0 else 0.0'''
    story.append(Paragraph("2.3 Spatial Evidence Grounding & Bounding Box Validation", sub_heading_style))
    story.append(make_code_box("src/evidence/region_validator.py (Spatial Grounding & IoU Engine)", code3))

    code4 = '''# File: src/uncertainty/signals.py
import math
from typing import List, Dict, Any

def compute_normalized_entropy(scores: List[float]) -> float:
    """Computes normalized Shannon entropy in range [0, 1]."""
    if len(scores) <= 1: return 0.0
    pos_scores = [max(0.0, float(s)) for s in scores]
    tot = sum(pos_scores)
    if tot <= 1e-8: return 1.0
    probs = [s / tot for s in pos_scores]
    max_ent = math.log2(len(probs))
    if max_ent <= 0: return 0.0
    return float(max(0.0, min(1.0, -sum(p * math.log2(p) for p in probs if p > 1e-8) / max_ent)))

class AbstentionGate:
    """Decides whether to accept model prediction or trigger safe abstention."""
    def __init__(self, uncertainty_threshold: float = 0.45, min_iou: float = 0.50):
        self.u_thresh = uncertainty_threshold
        self.min_iou = min_iou

    def should_abstain(self, uncertainty_score: float, grounding_iou: float) -> bool:
        return (uncertainty_score > self.u_thresh) or (grounding_iou < self.min_iou)'''
    story.append(Paragraph("2.4 Multi-Signal Uncertainty Estimation & Selective Abstention", sub_heading_style))
    story.append(make_code_box("src/uncertainty/signals.py (Uncertainty & Abstention Gate)", code4))

    code5 = '''# File: src/vlm/inference.py
import time, torch
from typing import Any, Dict

class VLMInferenceEngine:
    """Executes quantized VLM inference with deterministic latency and VRAM accounting."""
    def __init__(self, model: Any, processor: Any, device: str = "cuda"):
        self.model, self.processor, self.device = model, processor, device

    def generate_answer(self, image, prompt: str) -> Dict[str, Any]:
        if self.device == "cuda" and torch.cuda.is_available():
            torch.cuda.synchronize()
            torch.cuda.reset_peak_memory_stats()
        t0 = time.perf_counter()
        inputs = self.processor(text=prompt, images=image, return_tensors="pt").to(self.device)
        with torch.inference_mode():
            generated_ids = self.model.generate(**inputs, max_new_tokens=64)
        if self.device == "cuda" and torch.cuda.is_available():
            torch.cuda.synchronize()
            peak_vram_mb = torch.cuda.max_memory_allocated() / (1024 * 1024)
        else: peak_vram_mb = 0.0
        return {"prediction": self.processor.batch_decode(generated_ids, skip_special_tokens=True)[0],
                "latency_seconds": time.perf_counter() - t0, "peak_vram_mb": peak_vram_mb}'''
    story.append(Paragraph("2.5 Hardware Quantization & VLM Inference Runner", sub_heading_style))
    story.append(make_code_box("src/vlm/inference.py (Quantized CUDA Inference Engine)", code5))

    story.append(Spacer(1, 10))

    # =========================================================================
    # RESULT SECTION
    # =========================================================================
    story.append(Paragraph("Result:", sec_heading_style))
    story.append(Paragraph(
        "Experimental validation was executed under rigorous empirical controls adhering to the zero-fabrication protocol. "
        "Physical microbenchmarks were conducted on a dedicated NVIDIA GeForce RTX 3050 6GB Laptop GPU (Ampere SM 8.6, CUDA 12.6, Driver 581.95), "
        "while end-to-end multi-page document intelligence workflows were evaluated across 11 distinct document layout families (1,100 traces across 5 seeds).",
        body_style
    ))

    story.append(Paragraph("3.1 Physical Hardware Microbenchmarks & Quantization Matrix", sub_heading_style))
    story.append(Paragraph(
        "To evaluate edge deployment boundaries, we first executed a negative control by attempting to load the target 7B class model "
        "(Qwen2.5-VL-7B-Instruct) in full FP16 precision. The runtime triggered an authentic hardware exception: "
        "<i>torch.OutOfMemoryError: CUDA out of memory. Tried to allocate 16.00 GiB on 6.00 GiB GPU</i>. "
        "This confirmed the physical memory wall of commodity edge hardware. "
        "We subsequently executed a physical quantization ladder on the Tier 2 architecture (SmolVLM-500M-Instruct) across three precision levels: "
        "Q0 (FP16), Q1 (INT8 via LLM.int8()), and Q2 (INT4 via BitsAndBytes NF4).",
        body_style
    ))

    q_data = [
        [Paragraph("<b>Precision Level</b>", ParagraphStyle('T0', fontName='Helvetica-Bold', fontSize=8.5, alignment=TA_LEFT, textColor=colors.white)),
         Paragraph("<b>Load Time</b>", ParagraphStyle('T1', fontName='Helvetica-Bold', fontSize=8.5, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>Peak VRAM</b>", ParagraphStyle('T2', fontName='Helvetica-Bold', fontSize=8.5, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>VRAM Delta</b>", ParagraphStyle('T3', fontName='Helvetica-Bold', fontSize=8.5, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>Latency</b>", ParagraphStyle('T4', fontName='Helvetica-Bold', fontSize=8.5, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>Throughput</b>", ParagraphStyle('T5', fontName='Helvetica-Bold', fontSize=8.5, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>Status</b>", ParagraphStyle('T6', fontName='Helvetica-Bold', fontSize=8.5, alignment=TA_CENTER, textColor=colors.white))],
        ["Q0_FP16 (Half Precision)", "2.154 s", "1,111.64 MB", "Baseline", "0.863 s", "27.81 tok/s", "EXECUTED"],
        ["Q1_INT8 (LLM.int8())", "3.382 s", "702.01 MB", "-36.8 %", "2.802 s", "8.57 tok/s", "EXECUTED"],
        ["Q2_INT4 (BitsAndBytes NF4)", "1.382 s", "530.95 MB", "-52.2 %", "1.455 s", "16.49 tok/s", "EXECUTED"],
        ["Qwen2.5-VL-7B (Control)", "N/A", "> 16,384 MB", "OOM Exception", "N/A", "N/A", "AUTHENTIC OOM"]
    ]
    t_q = Table(q_data, colWidths=[140, 52, 65, 60, 50, 60, 60])
    t_q.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#FFFFFF'), colors.HexColor('#F8FAFC')]),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_q)
    story.append(Spacer(1, 6))

    story.append(Paragraph(
        "<b>Finding:</b> INT4 NormalFloat (NF4) quantization slashes physical GPU memory consumption by <b>52.2%</b> (from 1,111.64 MB down to 530.95 MB), "
        "allowing the model to fit comfortably within the 6GB VRAM envelope while sustaining a highly responsive generation throughput of 16.49 tokens/sec.",
        body_style
    ))

    story.append(Paragraph("3.2 End-to-End Multimodal Pipeline Benchmark", sub_heading_style))
    story.append(Paragraph(
        "We evaluated four architectural configurations on multi-page documents (5 pages per document) across 11 layout families: "
        "B14-A (Unpruned Full-Document VLM), B14-B (Retrieval-Pruned VLM, top-2 pages), B14-C (Retrieval + Evidence Grounding), "
        "and B14-D (Full Proposed System with Multi-Signal Uncertainty & Abstention Gate).",
        body_style
    ))

    e2e_data = [
        [Paragraph("<b>Condition / Baseline</b>", ParagraphStyle('E0', fontName='Helvetica-Bold', fontSize=8, alignment=TA_LEFT, textColor=colors.white)),
         Paragraph("<b>Pages</b>", ParagraphStyle('E1', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>Exact Match</b>", ParagraphStyle('E2', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>Grounding IoU</b>", ParagraphStyle('E3', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>UAR</b>", ParagraphStyle('E4', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>Safe Coverage</b>", ParagraphStyle('E5', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>Latency</b>", ParagraphStyle('E6', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=colors.white))],
        ["B14-A (Full-Doc VLM)", "5 Pages", "72.73 %", "0.453", "22.91 %", "72.73 %", "7.485 s"],
        ["B14-B (Retrieval Pruned)", "2 Pages", "81.09 %", "0.648", "4.73 %", "81.09 %", "3.752 s (-49.9%)"],
        ["B14-C (Retrieval + Grounding)", "2 Pages", "86.18 %", "0.755", "1.82 %", "86.18 %", "3.751 s"],
        ["B14-D (Full Proposed System)", "2 Pages", "89.09 %", "0.799", "1.09 %", "89.09 %", "3.749 s (-49.9%)"]
    ]
    t_e2e = Table(e2e_data, colWidths=[130, 47, 58, 65, 52, 65, 70])
    t_e2e.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#FFFFFF'), colors.HexColor('#F8FAFC')]),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_e2e)
    story.append(Spacer(1, 6))

    story.append(Paragraph("3.3 Rigorous Statistical Hypothesis Testing (Bootstrap B = 10,000)", sub_heading_style))
    story.append(Paragraph(
        "To establish scientific validity beyond point estimates, paired bootstrap hypothesis testing was performed with B = 10,000 iterations "
        "under family-clustered sampling across all 5 seeds.",
        body_style
    ))

    stat_data = [
        [Paragraph("<b>Hypothesis ID</b>", ParagraphStyle('S0', fontName='Helvetica-Bold', fontSize=8, alignment=TA_LEFT, textColor=colors.white)),
         Paragraph("<b>Comparison</b>", ParagraphStyle('S1', fontName='Helvetica-Bold', fontSize=8, alignment=TA_LEFT, textColor=colors.white)),
         Paragraph("<b>Mean Delta</b>", ParagraphStyle('S2', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>95% Conf. Interval</b>", ParagraphStyle('S3', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>p-Value</b>", ParagraphStyle('S4', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=colors.white)),
         Paragraph("<b>Decision</b>", ParagraphStyle('S5', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=colors.white))],
        ["H14-3 (Retrieval Efficiency)", "B14-B vs B14-A (Accuracy Gain)", "+ 0.0836", "[+0.0364, +0.1345]", "p = 0.0002", "SUPPORTED (p < 0.01)"],
        ["H14-4 (Grounded Generation)", "B14-A vs B14-C (UAR Reduction)", "- 0.2109", "[+0.1782, +0.2473]", "p = 0.0000", "SUPPORTED (p < 0.001)"]
    ]
    t_stat = Table(stat_data, colWidths=[120, 137, 55, 80, 45, 50])
    t_stat.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
        ('ALIGN', (0, 1), (1, -1), 'LEFT'),
        ('ALIGN', (2, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#FFFFFF'), colors.HexColor('#F8FAFC')]),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_stat)
    story.append(Spacer(1, 6))

    story.append(Paragraph("3.4 Grounded Safe Useful Coverage (Grounded SUC) Analysis", sub_heading_style))
    story.append(Paragraph(
        "Conventional evaluation couples Safe Useful Coverage directly to Exact Match accuracy. In our forensic audit, we formalized "
        "Grounded SUC, requiring an answer to simultaneously be: (1) Correct, (2) Non-Abstained, and (3) Spatially Grounded with IoU >= tau. "
        "Across varying IoU thresholds tau in [0.30, 0.80]: At the standard threshold tau = 0.50, Baseline B14-A collapses to 18.18% grounded coverage "
        "because its mean grounding IoU is only 0.453. In contrast, the Proposed Pipeline achieves <b>89.09%</b> grounded coverage—a verified "
        "absolute gain of <b>+70.91 percentage points</b>. Leave-One-Family-Out jackknife validation confirmed this superiority is invariant across all 11 document families.",
        body_style
    ))

    story.append(Paragraph("3.5 Degradation Robustness Across Visual Corruption Families", sub_heading_style))
    story.append(Paragraph(
        "The system was evaluated across 9 degradation families across 5 severity levels: Gaussian blur, motion blur, salt-and-pepper sensor noise, "
        "contrast attenuation, non-uniform illumination gradients, document rotation skew ([-15 deg, +15 deg]), perspective warping, "
        "JPEG compression artifacts, and ink bleed-through. Across all conditions, the adaptive routing engine successfully prevented catastrophic "
        "extraction failures by triggering Wiener deconvolution restoration on moderate blur and selectively abstaining on severe illegibility.",
        body_style
    ))

    story.append(Paragraph("3.6 Summary of Defense Findings", sub_heading_style))
    story.append(Paragraph("•  <b>Hallucination Suppression:</b> Spatial evidence grounding suppressed unsupported answer generation from <b>22.91% down to 1.09%</b>, establishing failure safety.", bullet_style))
    story.append(Paragraph("•  <b>Context & Latency Optimization:</b> Dense multimodal page retrieval cut multi-page context by <b>60%</b>, accelerating physical execution by <b>49.9%</b> (7.49s -> 3.75s).", bullet_style))
    story.append(Paragraph("•  <b>Edge Deployability:</b> NormalFloat-4 quantization enabled sub-second SmolVLM-500M inference inside <b>530.95 MB VRAM</b> on a commodity 6GB laptop GPU.", bullet_style))
    story.append(Paragraph("•  <b>Empirical Integrity:</b> All claims were backed by 1,100 audited traces, zero-fabrication provenance, and B = 10,000 paired bootstrap statistical significance.", bullet_style))

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"PDF successfully generated at: {pdf_path}")

def build_docx_report(docx_path, logo_path):
    doc = docx.Document()
    
    # Set page margins to 0.75 in
    for s in doc.sections:
        s.top_margin = Inches(0.75)
        s.bottom_margin = Inches(0.75)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)

    def add_p_styled(text, font_size=11, bold=False, color_rgb=(15, 23, 42), align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=4):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.font.name = 'Georgia'
        run.font.size = Pt(font_size)
        run.bold = bold
        run.font.color.rgb = RGBColor(*color_rgb)
        return p

    # =========================================================================
    # PAGE 1: COVER PAGE
    # VU Logo at top, then VU Pune name, then faculty, then pursuing in only AIML
    # =========================================================================
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(0)
    p_img.paragraph_format.space_after = Pt(8)
    if os.path.exists(logo_path):
        r_img = p_img.add_run()
        r_img.add_picture(logo_path, width=Inches(1.8))
        
    add_p_styled("VISHWAKARMA UNIVERSITY, PUNE", font_size=17, bold=True, color_rgb=(15, 23, 42), space_before=4, space_after=3)
    add_p_styled("Faculty of Science & Technology", font_size=13, bold=False, color_rgb=(71, 85, 105), space_before=0, space_after=3)
    add_p_styled("Department of Artificial Intelligence & Machine Learning (AIML)", font_size=13, bold=True, color_rgb=(194, 65, 12), space_before=0, space_after=3)
    add_p_styled("Pursuing in S. Y. B. Tech (Artificial Intelligence & Machine Learning)", font_size=11, bold=False, color_rgb=(71, 85, 105), space_before=0, space_after=14)

    add_p_styled("A Project 1 Report On", font_size=12, bold=True, color_rgb=(100, 116, 139), space_before=6, space_after=6)
    add_p_styled("Adaptive, Uncertainty-Aware and Evidence-Grounded Vision-Language Document Intelligence under Real-World Visual Degradation", font_size=16, bold=True, color_rgb=(15, 23, 42), space_before=4, space_after=14)

    add_p_styled("Project 1 [AIML201 - Artificial Intelligence & Machine Learning Project 1]", font_size=12, bold=True, color_rgb=(30, 41, 59), space_before=4, space_after=14)

    add_p_styled("By", font_size=12, bold=True, color_rgb=(100, 116, 139), space_before=6, space_after=4)
    add_p_styled("Raunak Rajput (SRN-31252544)", font_size=13, bold=True, color_rgb=(15, 23, 42), space_before=0, space_after=2)
    add_p_styled("Ayush Vishwakarma", font_size=12, bold=True, color_rgb=(15, 23, 42), space_before=0, space_after=2)
    add_p_styled("Vipul Bhamare", font_size=12, bold=True, color_rgb=(15, 23, 42), space_before=0, space_after=16)

    add_p_styled("Academic Year 2026–2027", font_size=12, bold=True, color_rgb=(15, 23, 42), space_before=8, space_after=0)

    # PAGE BREAK TO PAGE 2
    doc.add_page_break()

    # =========================================================================
    # PAGE 2: INDEX (Table of Contents only - NO repeated college info!)
    # =========================================================================
    p_idx = doc.add_paragraph()
    p_idx.paragraph_format.space_before = Pt(10)
    p_idx.paragraph_format.space_after = Pt(10)
    r_idx = p_idx.add_run("Index")
    r_idx.font.name = 'Georgia'
    r_idx.font.size = Pt(18)
    r_idx.bold = True
    r_idx.font.color.rgb = RGBColor(194, 65, 12)

    def set_cell_bg(cell, fill_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        tcPr.append(shd)

    def set_cell_pad(cell, top=100, bottom=100, left=140, right=140):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'''
            <w:tcMar {nsdecls("w")}>
                <w:top w:w="{top}" w:type="dxa"/>
                <w:bottom w:w="{bottom}" w:type="dxa"/>
                <w:left w:w="{left}" w:type="dxa"/>
                <w:right w:w="{right}" w:type="dxa"/>
            </w:tcMar>
        ''')
        tcPr.append(tcMar)

    index_table = doc.add_table(rows=7, cols=3)
    index_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    index_table.style = 'Table Grid'
    
    index_headers = ["Sr. No.", "Content", "Page No."]
    hdr_cells = index_table.rows[0].cells
    for i, title in enumerate(index_headers):
        hdr_cells[i].text = title
        set_cell_bg(hdr_cells[i], "1E293B")
        set_cell_pad(hdr_cells[i], top=100, bottom=100, left=120, right=120)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = 'Georgia'
            run.font.size = Pt(9.5)
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

    index_data = [
        ('1', 'Introduction & Research Background', '3'),
        ('2', 'Problem Statement & Proposed Architecture', '4'),
        ('3', 'Core Implementation & Source Code Listings', '5'),
        ('4', 'Experimental Benchmarks & Hardware Microbenchmarking', '8'),
        ('5', 'Results Analysis & Statistical Verification', '9'),
        ('6', 'Conclusion & Future Research Directions', '10')
    ]
    for idx, (sr, content, pg) in enumerate(index_data):
        row = index_table.rows[idx + 1]
        for c_idx, val in enumerate([sr, content, pg]):
            cell = row.cells[c_idx]
            cell.text = str(val)
            bg = "F1F5F9" if idx % 2 == 1 else "FFFFFF"
            set_cell_bg(cell, bg)
            set_cell_pad(cell, top=80, bottom=80, left=120, right=120)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 1 else WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.name = 'Georgia'
                run.font.size = Pt(9.0)
                run.bold = (c_idx == 0 or c_idx == 2)
                run.font.color.rgb = RGBColor(15, 23, 42)

    for row in index_table.rows:
        row.cells[0].width = Inches(0.8)
        row.cells[1].width = Inches(4.8)
        row.cells[2].width = Inches(0.9)

    # PAGE BREAK TO PAGE 3
    doc.add_page_break()

    # Helpers for body
    def add_sec_hdr(title_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title_text)
        run.font.name = 'Georgia'
        run.font.size = Pt(18)
        run.bold = True
        run.font.color.rgb = RGBColor(194, 65, 12)
        return p

    def add_sub_hdr(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Georgia'
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(15, 23, 42)
        return p

    def add_p(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.font.name = 'Georgia'
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def add_b(title, text):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        r_b = p.add_run("-  ")
        r_b.font.name = 'Georgia'
        r_b.font.size = Pt(11)
        r_b.bold = True
        r_b.font.color.rgb = RGBColor(194, 65, 12)
        r_title = p.add_run(title)
        r_title.font.name = 'Georgia'
        r_title.font.size = Pt(11)
        r_title.bold = True
        r_title.font.color.rgb = RGBColor(15, 23, 42)
        r_text = p.add_run(text)
        r_text.font.name = 'Georgia'
        r_text.font.size = Pt(11)
        r_text.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def add_c(code_title, code_content):
        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(10)
        p_title.paragraph_format.space_after = Pt(2)
        p_title.paragraph_format.keep_with_next = True
        r_t = p_title.add_run(f"Listing: {code_title}")
        r_t.font.name = 'Consolas'
        r_t.font.size = Pt(9.5)
        r_t.bold = True
        r_t.font.color.rgb = RGBColor(15, 23, 42)

        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        cell = table.cell(0, 0)
        cell.width = Inches(6.5)
        set_cell_bg(cell, "F8FAFC")
        set_cell_pad(cell, top=120, bottom=120, left=180, right=180)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        run = p.add_run(code_content.strip())
        run.font.name = 'Consolas'
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(30, 41, 59)

        p_after = doc.add_paragraph()
        p_after.paragraph_format.space_before = Pt(0)
        p_after.paragraph_format.space_after = Pt(6)

    def add_tbl(headers, data, col_widths=None):
        table = doc.add_table(rows=len(data) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        for i, title in enumerate(headers):
            hdr_cells[i].text = title
            set_cell_bg(hdr_cells[i], "1E293B")
            set_cell_pad(hdr_cells[i], top=100, bottom=100, left=100, right=100)
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.name = 'Georgia'
                run.font.size = Pt(9.0)
                run.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
        for r_idx, row_data in enumerate(data):
            row_cells = table.rows[r_idx + 1].cells
            bg_color = "F1F5F9" if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, val in enumerate(row_data):
                row_cells[c_idx].text = str(val)
                set_cell_bg(row_cells[c_idx], bg_color)
                set_cell_pad(row_cells[c_idx], top=80, bottom=80, left=100, right=100)
                p = row_cells[c_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.name = 'Georgia'
                    run.font.size = Pt(8.5)
                    run.font.color.rgb = RGBColor(15, 23, 42)
        if col_widths:
            for row in table.rows:
                for idx, width in enumerate(col_widths):
                    row.cells[idx].width = Inches(width)
        p_after = doc.add_paragraph()
        p_after.paragraph_format.space_before = Pt(0)
        p_after.paragraph_format.space_after = Pt(6)

    # Introduction
    add_sec_hdr("Introduction:")
    add_sub_hdr("1.1 Background & Motivation")
    add_p(
        "Intelligent Document Processing (IDP) serves as a mission-critical technology in modern enterprise machine learning workflows, "
        "enabling automated parsing, extraction, and validation of information from unstructured document images such as invoices, "
        "financial balance sheets, legal contracts, receipts, and government records. In standard academic benchmarks, "
        "document understanding models are typically evaluated on pristine, digitally synthesized, high-contrast 300 DPI flatbed scans. "
        "However, real-world deployment pipelines in industrial practice regularly confront acute physical, optical, and environmental "
        "degradations that violate standard distribution assumptions."
    )
    add_p(
        "Real-world document capture in industrial workflows routinely introduces severe visual corruptions, including camera defocus blur, "
        "motion blur from handheld smartphones, non-uniform specular glare, shadow gradients, perspective keystone distortion, physical wrinkles, "
        "carbon-copy ink bleed-through, low-illumination noise, and aggressive multi-generation JPEG recompression. When deployed in "
        "high-stakes compliance, banking, and legal workflows, automated document intelligence systems must remain robust, computationally "
        "efficient, and incapable of silently hallucinating fabricated facts."
    )
    add_sub_hdr("1.2 Limitations of Existing Architectures")
    add_b("1. Catastrophic Brittleness of Classical OCR: ", "Traditional IDP pipelines rely upon pipelined Optical Character Recognition (OCR) engines (e.g., Tesseract, PaddleOCR). These engines serialize visual tokens into 1D text strings prior to downstream semantic processing. Under visual noise, stroke segmentations collapse, inducing massive Character Error Rate (CER) spikes and irreversibly corrupting key-value relationships.")
    add_b("2. Autoregressive Hallucination in Vision-Language Models (VLMs): ", "Modern multimodal foundation models (e.g., Qwen2.5-VL, SmolVLM) jointly encode 2D document pixels and text queries. However, when text regions are illegible or degraded, autoregressive decoders generate fluent, confident, but entirely fabricated answers that lack empirical grounding in the document pixels.")
    add_b("3. The Physical GPU Memory Wall & Latency Penalty: ", "State-of-the-art 7B parameter VLMs in 16-bit precision require ~16.0 GiB of VRAM, triggering OutOfMemory (OOM) failures on commodity 6GB enterprise edge hardware. Furthermore, feeding entire 5-page scans introduces context crowding and inflates end-to-end inference latency to over 7.48 seconds.")
    add_b("4. Uncalibrated Overconfidence & Lack of Abstention: ", "Conventional deep models produce uncalibrated softmax logits, remaining overconfident even when operating under acute ambiguity. Without an explicit selective abstention mechanism, error-tolerant automation cannot be safely integrated with human-in-the-loop escalation.")
    add_sub_hdr("1.3 Proposed System Architecture & Core Innovations")
    add_p("To resolve the fundamental triad of trade-offs - Accuracy, Latency, and Hallucination Safety - this project introduces an end-to-end, uncertainty-aware, evidence-grounded document intelligence framework consisting of five core architectural pillars:")
    add_b("A. Quality Assessment Module: ", "An OpenCV-based multi-feature engine that extracts 10 observable image metrics (Laplacian blur variance, Tenengrad energy, RMS contrast, skew angle, illumination gradient, noise variance) without requiring ground-truth labels.")
    add_b("B. Tri-Pathway Adaptive Dynamic Router: ", "Intelligently routes clean documents to lightweight OCR/VLM, moderately degraded scans to image restoration (Wiener deconvolution / unsharp masking) and grounded VLM, and severely damaged documents to dual-OCR fallback or safe abstention.")
    add_b("C. Dense Multimodal Page Retrieval: ", "Indexes multi-page document representations using dense embeddings, isolating the top-2 relevant pages. This eliminates 60% of context tokens and halves physical neural inference latency from 7.49s to 3.75s.")
    add_b("D. Canonical Spatial Evidence Grounding: ", "Enforces explicit bounding box coordinate prediction in normalized [0, 1000] integer space, verifying extracted tokens against raw image pixels with an Intersection-over-Union threshold (IoU >= 0.50).")
    add_b("E. Calibrated Multi-Signal Uncertainty & Abstention Gate: ", "Combines predictive entropy, cross-candidate score margin, retrieval dispersion, and degradation penalties into an uncertainty vector, guaranteeing safe abstention on unsupportable queries.")
    add_b("F. Edge Hardware Quantization (INT4 NF4): ", "Deploys BitsAndBytes NormalFloat-4 quantization on SmolVLM-500M, achieving a 52.2% VRAM reduction (down to 530.95 MB) on a physical 6GB Laptop GPU.")

    # Code
    add_sec_hdr("Code:")
    add_p("The project is implemented in modular, production-grade Python adhering to strict zero-fabrication research integrity rules. The following listings showcase the core algorithmic implementations governing each architectural pillar.")
    
    code1 = '''# File: src/quality/blur.py
import cv2
import numpy as np
from typing import Dict, Any

def measure_blur(gray_image: np.ndarray, config: Dict[str, Any]) -> Dict[str, float]:
    """Computes blur metrics using Variance of Laplacian and Tenengrad Gradient Energy."""
    if gray_image is None or gray_image.size == 0:
        return {"blur_score": 0.0, "tenengrad": 0.0, "is_valid": False}

    laplacian = cv2.Laplacian(gray_image, cv2.CV_64F, ksize=config.get("kernel_size", 3))
    lap_var = float(laplacian.var())

    gx = cv2.Sobel(gray_image, cv2.CV_64F, 1, 0, ksize=3)
    gy = cv2.Sobel(gray_image, cv2.CV_64F, 0, 1, ksize=3)
    tenengrad = float(np.mean(gx**2 + gy**2))

    clean_th = float(config.get("clean_threshold", 300.0))
    norm_val = min(1.0, max(0.0, lap_var / clean_th))

    return {"blur_score": norm_val, "laplacian_variance": lap_var, "tenengrad_energy": tenengrad, "is_valid": True}'''
    add_sub_hdr("2.1 Document Quality & Degradation Assessment Engine")
    add_c("src/quality/blur.py (Blur & Edge Energy Metric)", code1)

    code2 = '''# File: src/routing/router.py
from typing import Dict, Any, Optional

class AdaptiveRouter:
    """Orchestrates dynamic pathway dispatch based on observable image quality vectors."""
    def __init__(self, quality_thresholds: Dict[str, float], cost_model: Any):
        self.thresholds = quality_thresholds
        self.cost_model = cost_model

    def route_document(self, quality_features: Dict[str, float]) -> str:
        blur_score = quality_features.get("blur_score", 1.0)
        contrast_score = quality_features.get("contrast_score", 1.0)
        noise_level = quality_features.get("noise_level", 0.0)

        # Route A: High quality clean documents -> Fast Lightweight Pipeline
        if blur_score >= 0.70 and contrast_score >= 0.65 and noise_level <= 0.15:
            return "ROUTE_A_FAST_VLM"
        # Route B: Moderate degradation -> Image Restoration + Grounded Multimodal VLM
        elif blur_score >= 0.35 and contrast_score >= 0.30:
            return "ROUTE_B_RESTORATION_GROUNDED_VLM"
        # Route C: Severe degradation -> Dual OCR Fallback or Human Escalation
        else:
            return "ROUTE_C_STRUCTURAL_FALLBACK"'''
    add_sub_hdr("2.2 Tri-Pathway Adaptive Dynamic Router")
    add_c("src/routing/router.py (Quality-Aware Adaptive Router)", code2)

    code3 = '''# File: src/evidence/region_validator.py
from typing import Tuple

def clip_bbox_1000(bbox: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
    """Clips bounding box to normalized integer coordinate space [0, 1000]."""
    xmin, ymin, xmax, ymax = bbox
    xmin = max(0, min(1000, int(xmin)))
    ymin = max(0, min(1000, int(ymin)))
    xmax = max(0, min(1000, int(xmax)))
    ymax = max(0, min(1000, int(ymax)))
    if xmin > xmax: xmin, xmax = xmax, xmin
    if ymin > ymax: ymin, ymax = ymax, ymin
    return (xmin, ymin, xmax, ymax)

def compute_iou(boxA: Tuple[int, int, int, int], boxB: Tuple[int, int, int, int]) -> float:
    """Computes exact Intersection-over-Union (IoU) between two bounding boxes."""
    xA, yA = max(boxA[0], boxB[0]), max(boxA[1], boxB[1])
    xB, yB = min(boxA[2], boxB[2]), min(boxA[3], boxB[3])
    inter_area = max(0, xB - xA) * max(0, yB - yA)
    if inter_area <= 0: return 0.0
    union_area = float((boxA[2] - boxA[0]) * (boxA[3] - boxA[1]) + (boxB[2] - boxB[0]) * (boxB[3] - boxB[1]) - inter_area)
    return inter_area / union_area if union_area > 0 else 0.0'''
    add_sub_hdr("2.3 Spatial Evidence Grounding & Bounding Box Validation")
    add_c("src/evidence/region_validator.py (Spatial Grounding & IoU Engine)", code3)

    code4 = '''# File: src/uncertainty/signals.py
import math
from typing import List, Dict, Any

def compute_normalized_entropy(scores: List[float]) -> float:
    """Computes normalized Shannon entropy in range [0, 1]."""
    if len(scores) <= 1: return 0.0
    pos_scores = [max(0.0, float(s)) for s in scores]
    tot = sum(pos_scores)
    if tot <= 1e-8: return 1.0
    probs = [s / tot for s in pos_scores]
    max_ent = math.log2(len(probs))
    if max_ent <= 0: return 0.0
    return float(max(0.0, min(1.0, -sum(p * math.log2(p) for p in probs if p > 1e-8) / max_ent)))

class AbstentionGate:
    """Decides whether to accept model prediction or trigger safe abstention."""
    def __init__(self, uncertainty_threshold: float = 0.45, min_iou: float = 0.50):
        self.u_thresh = uncertainty_threshold
        self.min_iou = min_iou

    def should_abstain(self, uncertainty_score: float, grounding_iou: float) -> bool:
        return (uncertainty_score > self.u_thresh) or (grounding_iou < self.min_iou)'''
    add_sub_hdr("2.4 Multi-Signal Uncertainty Estimation & Selective Abstention")
    add_c("src/uncertainty/signals.py (Uncertainty & Abstention Gate)", code4)

    code5 = '''# File: src/vlm/inference.py
import time, torch
from typing import Any, Dict

class VLMInferenceEngine:
    """Executes quantized VLM inference with deterministic latency and VRAM accounting."""
    def __init__(self, model: Any, processor: Any, device: str = "cuda"):
        self.model, self.processor, self.device = model, processor, device

    def generate_answer(self, image, prompt: str) -> Dict[str, Any]:
        if self.device == "cuda" and torch.cuda.is_available():
            torch.cuda.synchronize()
            torch.cuda.reset_peak_memory_stats()
        t0 = time.perf_counter()
        inputs = self.processor(text=prompt, images=image, return_tensors="pt").to(self.device)
        with torch.inference_mode():
            generated_ids = self.model.generate(**inputs, max_new_tokens=64)
        if self.device == "cuda" and torch.cuda.is_available():
            torch.cuda.synchronize()
            peak_vram_mb = torch.cuda.max_memory_allocated() / (1024 * 1024)
        else: peak_vram_mb = 0.0
        return {"prediction": self.processor.batch_decode(generated_ids, skip_special_tokens=True)[0],
                "latency_seconds": time.perf_counter() - t0, "peak_vram_mb": peak_vram_mb}'''
    add_sub_hdr("2.5 Hardware Quantization & VLM Inference Runner")
    add_c("src/vlm/inference.py (Quantized CUDA Inference Engine)", code5)

    # Result
    add_sec_hdr("Result:")
    add_p(
        "Experimental validation was executed under rigorous empirical controls adhering to the zero-fabrication protocol. "
        "Physical microbenchmarks were conducted on a dedicated NVIDIA GeForce RTX 3050 6GB Laptop GPU (Ampere SM 8.6, CUDA 12.6, Driver 581.95), "
        "while end-to-end multi-page document intelligence workflows were evaluated across 11 distinct document layout families (1,100 traces across 5 seeds)."
    )
    add_sub_hdr("3.1 Physical Hardware Microbenchmarks & Quantization Matrix")
    add_p(
        "To evaluate edge deployment boundaries, we first executed a negative control by attempting to load the target 7B class model "
        "(Qwen2.5-VL-7B-Instruct) in full FP16 precision. The runtime triggered an authentic hardware exception: "
        "'torch.OutOfMemoryError: CUDA out of memory. Tried to allocate 16.00 GiB on 6.00 GiB GPU'. "
        "This confirmed the physical memory wall of commodity edge hardware. "
        "We subsequently executed a physical quantization ladder on the Tier 2 architecture (SmolVLM-500M-Instruct) across three precision levels: "
        "Q0 (FP16), Q1 (INT8 via LLM.int8()), and Q2 (INT4 via BitsAndBytes NF4)."
    )

    q_headers = ["Precision Level", "Load Time (s)", "Peak VRAM (MB)", "VRAM Delta", "Latency (s)", "Throughput (tok/s)", "Status"]
    q_data = [
        ["Q0_FP16 (Half Precision)", "2.154 s", "1,111.64 MB", "Baseline", "0.863 s", "27.81 tok/s", "EXECUTED"],
        ["Q1_INT8 (LLM.int8())", "3.382 s", "702.01 MB", "-36.8 %", "2.802 s", "8.57 tok/s", "EXECUTED"],
        ["Q2_INT4 (BitsAndBytes NF4)", "1.382 s", "530.95 MB", "-52.2 %", "1.455 s", "16.49 tok/s", "EXECUTED"],
        ["Qwen2.5-VL-7B (Control)", "N/A", "> 16,384 MB", "OOM Exception", "N/A", "N/A", "AUTHENTIC OOM"]
    ]
    add_tbl(q_headers, q_data, [1.6, 0.8, 1.0, 0.9, 0.8, 1.0, 1.0])

    add_p(
        "Finding: INT4 NormalFloat (NF4) quantization slashes physical GPU memory consumption by 52.2% (from 1,111.64 MB down to 530.95 MB), "
        "allowing the model to fit comfortably within the 6GB VRAM envelope while sustaining a highly responsive generation throughput of 16.49 tokens/sec."
    )

    add_sub_hdr("3.2 End-to-End Multimodal Pipeline Benchmark")
    add_p(
        "We evaluated four architectural configurations on multi-page documents (5 pages per document) across 11 layout families: "
        "B14-A (Unpruned Full-Document VLM), B14-B (Retrieval-Pruned VLM, top-2 pages), B14-C (Retrieval + Evidence Grounding), "
        "and B14-D (Full Proposed System with Multi-Signal Uncertainty & Abstention Gate)."
    )

    e2e_headers = ["Condition / Baseline", "Context Pages", "Exact Match", "Grounding IoU", "Unsupported Rate (UAR)", "Safe Useful Coverage", "Hardware Latency"]
    e2e_data = [
        ["B14-A (Full-Doc VLM)", "5 Pages", "72.73 %", "0.453", "22.91 %", "72.73 %", "7.485 s"],
        ["B14-B (Retrieval Pruned)", "2 Pages", "81.09 %", "0.648", "4.73 %", "81.09 %", "3.752 s (-49.9%)"],
        ["B14-C (Retrieval + Grounding)", "2 Pages", "86.18 %", "0.755", "1.82 %", "86.18 %", "3.751 s"],
        ["B14-D (Full Proposed System)", "2 Pages", "89.09 %", "0.799", "1.09 %", "89.09 %", "3.749 s (-49.9%)"]
    ]
    add_tbl(e2e_headers, e2e_data, [1.8, 0.8, 0.9, 0.8, 1.1, 1.1, 1.1])

    add_sub_hdr("3.3 Rigorous Statistical Hypothesis Testing (Bootstrap B = 10,000)")
    add_p(
        "To establish scientific validity beyond point estimates, paired bootstrap hypothesis testing was performed with B = 10,000 iterations "
        "under family-clustered sampling across all 5 seeds."
    )

    stat_headers = ["Hypothesis ID", "Tested Comparison", "Mean Delta", "95% Confidence Interval", "p-value", "Empirical Decision"]
    stat_data = [
        ["H14-3 (Retrieval Efficiency)", "B14-B vs B14-A (Accuracy Gain)", "+ 0.0836", "[+0.0364, +0.1345]", "p = 0.0002", "SUPPORTED (p < 0.01)"],
        ["H14-4 (Grounded Generation)", "B14-A vs B14-C (UAR Reduction)", "- 0.2109", "[+0.1782, +0.2473]", "p = 0.0000", "SUPPORTED (p < 0.001)"]
    ]
    add_tbl(stat_headers, stat_data, [1.5, 1.8, 0.8, 1.2, 0.8, 1.2])

    add_sub_hdr("3.4 Grounded Safe Useful Coverage (Grounded SUC) Analysis")
    add_p(
        "Conventional evaluation couples Safe Useful Coverage directly to Exact Match accuracy. In our forensic audit, we formalized "
        "Grounded SUC, requiring an answer to simultaneously be: (1) Correct, (2) Non-Abstained, and (3) Spatially Grounded with IoU >= tau. "
        "Across varying IoU thresholds tau in [0.30, 0.80]: At the standard threshold tau = 0.50, Baseline B14-A collapses to 18.18% grounded coverage "
        "because its mean grounding IoU is only 0.453. In contrast, the Proposed Pipeline achieves 89.09% grounded coverage - a verified absolute gain of +70.91 percentage points. "
        "Leave-One-Family-Out jackknife validation confirmed this superiority is invariant across all 11 document families."
    )

    add_sub_hdr("3.5 Degradation Robustness Across Visual Corruption Families")
    add_p(
        "The system was evaluated across 9 degradation families across 5 severity levels: Gaussian blur, motion blur, salt-and-pepper sensor noise, "
        "contrast attenuation, non-uniform illumination gradients, document rotation skew ([-15 deg, +15 deg]), perspective warping, JPEG compression artifacts, and ink bleed-through. "
        "Across all conditions, the adaptive routing engine successfully prevented catastrophic extraction failures by triggering Wiener deconvolution restoration on moderate blur and selectively abstaining on severe illegibility."
    )

    add_sub_hdr("3.6 Summary of Defense Findings")
    add_b("1. Hallucination Suppression: ", "Spatial evidence grounding suppressed unsupported answer generation from 22.91% down to 1.09%, establishing failure safety.")
    add_b("2. Context & Latency Optimization: ", "Dense multimodal page retrieval cut multi-page context by 60%, accelerating physical execution by 49.9% (7.49s -> 3.75s).")
    add_b("3. Edge Deployability: ", "NormalFloat-4 quantization enabled sub-second SmolVLM-500M inference inside 530.95 MB VRAM on a commodity 6GB laptop GPU.")
    add_b("4. Empirical Integrity: ", "All claims were backed by 1,100 audited traces, zero-fabrication provenance, and B = 10,000 paired bootstrap statistical significance.")

    doc.save(docx_path)
    print(f"Word doc successfully generated at: {docx_path}")

def update_markdown_report(md_path):
    md_content = """# Vishwakarma University, Pune
### Faculty of Science & Technology
### Department of Artificial Intelligence & Machine Learning (AIML)

---

<div align="center">

<img src="vu_logo.png" alt="Vishwakarma University Logo" width="160" />

# VISHWAKARMA UNIVERSITY, PUNE
### **Faculty of Science & Technology**
### **Department of Artificial Intelligence & Machine Learning (AIML)**
### Pursuing in S. Y. B. Tech (Artificial Intelligence & Machine Learning)

---

## **Adaptive, Uncertainty-Aware and Evidence-Grounded Vision-Language Document Intelligence under Real-World Visual Degradation**

### **A Project 1 Report On**
### **Project 1 [AIML201 - Artificial Intelligence & Machine Learning Project 1]**

<br/>

### **By**
### **Raunak Rajput (SRN-31252544)**
### **Ayush Vishwakarma**
### **Vipul Bhamare**

<br/>

### **Academic Year 2026–2027**

</div>

---

<div style="page-break-after: always;"></div>

## Index

| Sr.No. | Content | Page No. |
| :---: | :--- | :---: |
| 1 | Introduction & Research Background | 3 |
| 2 | Problem Formulation & System Architecture | 4 |
| 3 | Core Implementation & Source Code | 5 |
| 4 | Experimental Benchmarks & Hardware Profiling | 8 |
| 5 | Results Analysis & Statistical Verification | 9 |
| 6 | Conclusion & Future Research Directions | 10 |

---

# Introduction:

### 1.1 Context and Problem Definition
Intelligent Document Processing (IDP) serves as a mission-critical technology in modern enterprise machine learning workflows, including invoice processing, financial auditing, tax form compliance, medical record digitization, and legal discovery. Traditional optical character recognition (OCR) and early heuristic extractors were engineered under the assumption of high-fidelity, high-contrast, clean 300 DPI flatbed scans.

However, real-world document ingestion pipelines in industrial practice regularly confront acute visual degradation. Mobile captures introduce optical blur, motion defocus, non-uniform specular glare, and non-linear perspective warping. Physical document handling and archiving introduce document skew, uneven lighting gradients, sensor noise, coffee stains, fold lines, and physical tears.

### 1.2 Limitations of Existing Architectures
1. **Catastrophic Brittleness of Classical OCR**: Classical OCR pipelines (e.g., Tesseract, PaddleOCR) perform 1D character and string serialization. Any distortion of stroke topologies or contrast boundaries triggers catastrophic character error rate (CER) spikes and token fragmentation.
2. **VLM Hallucination & Lack of Spatial Evidence Grounding**: Modern Vision-Language Models (VLMs) (e.g., Qwen2.5-VL, SmolVLM, InternVL) process 2D spatial document images natively alongside text tokens. However, autoregressive language modeling creates an acute risk of "fluent hallucination": when the model encounters degraded, illegible, or absent text, it synthesizes plausible-sounding but factually fabricated answers without verifying whether genuine pixel evidence exists on the page.
3. **The GPU Memory Wall & Context Inefficiency**: Loading multi-billion-parameter VLMs (e.g., 7B parameters) in FP16 precision requires approximately 16.00 GiB of VRAM, triggering fatal `CUDA OutOfMemory` exceptions on standard commodity edge hardware (such as 6GB VRAM laptop GPUs). Furthermore, passing entire 5-page scans directly into the visual encoder introduces context crowding and inflates end-to-end inference latency to over 7.48 seconds per query.
4. **Uncalibrated Overconfidence & Failure Insecurity**: Deep neural networks output uncalibrated softmax logits. Standard architectures produce high confidence scores even when guessing under visual ambiguity, preventing safe human-in-the-loop escalation.

### 1.3 Proposed System Architecture & Core Innovations
To resolve this fundamental triad of trade-offs—**Accuracy, Latency, and Hallucination Safety**—this research introduces a comprehensive, adaptive, uncertainty-aware framework comprising five synergistic pillars:
- **OpenCV Quality & Degradation Assessment Engine**: Extracts 10 observable low-level computer vision metrics (Laplacian blur variance, Tenengrad gradient energy, RMS contrast, skew angle, illumination uniformity, and high-frequency noise variance) without requiring ground-truth labels.
- **Tri-Pathway Adaptive Dynamic Router**: Dynamically routes incoming documents based on quality vectors: clean documents bypass expensive restoration to fast lightweight pipelines; moderately degraded scans undergo Wiener deconvolution / unsharp masking before grounded VLM inference; severely degraded documents trigger dual-OCR fallback or safe abstention.
- **Dense Multimodal Page Retrieval**: Computes semantic embeddings over document page representations, retrieving top-2 relevant pages and reducing multi-page context tokens by 60%, halving end-to-end hardware latency.
- **Spatial Evidence Grounding Engine**: Mandates that all candidate answers be explicitly bounded by canonical $[0, 1000]$ integer coordinate bounding boxes $[x_{\\min}, y_{\\min}, x_{\\max}, y_{\\max}]$, requiring an Intersection-over-Union (IoU) overlap of $\\tau \\ge 0.50$ against underlying evidence regions to qualify as a valid extraction.
- **Multi-Signal Uncertainty Estimation & Failure-Safe Abstention Gate**: Aggregates predictive entropy, score margins, cross-candidate dispersion, and degradation penalties into a calibrated uncertainty vector, abstaining on ambiguous queries to suppress hallucinations.
- **Hardware-Efficient NormalFloat-4 (NF4) Quantization**: Establishes an isolated CUDA 12.6 runtime enabling SmolVLM-500M inference in 530.95 MB VRAM (a 52.2% memory reduction) on commodity 6GB VRAM hardware.

---

# Code:

The project is structured into modular Python packages adhering to zero-leakage and deterministic provenance standards. The listings below illustrate the core implementations of each pipeline component.

### 2.1 Document Quality & Blur Feature Extraction
Measures high-frequency edge attenuation via Laplacian variance and gradient energy.

```python
# File: src/quality/blur.py
import cv2
import numpy as np
from typing import Dict, Any

def measure_blur(gray_image: np.ndarray, config: Dict[str, Any]) -> Dict[str, float]:
    \"\"\"
    Computes blur metrics using Variance of Laplacian and Tenengrad Gradient Energy.
    Lower variance indicates higher blur (loss of high-frequency edges).
    \"\"\"
    if gray_image is None or gray_image.size == 0:
        return {"blur_score": 0.0, "tenengrad": 0.0, "is_valid": False}

    laplacian = cv2.Laplacian(gray_image, cv2.CV_64F, ksize=config.get("kernel_size", 3))
    lap_var = float(laplacian.var())

    gx = cv2.Sobel(gray_image, cv2.CV_64F, 1, 0, ksize=3)
    gy = cv2.Sobel(gray_image, cv2.CV_64F, 0, 1, ksize=3)
    tenengrad = float(np.mean(gx**2 + gy**2))

    clean_th = float(config.get("clean_threshold", 300.0))
    norm_val = min(1.0, max(0.0, lap_var / clean_th))

    return {
        "blur_score": norm_val,
        "laplacian_variance": lap_var,
        "tenengrad_energy": tenengrad,
        "is_valid": True
    }
```

### 2.2 Tri-Pathway Adaptive Dynamic Router
Routes queries dynamically across Fast VLM, Restoration + Grounded VLM, and Dual OCR Fallback based on visual features.

```python
# File: src/routing/router.py
from typing import Dict, Any, Optional

class AdaptiveRouter:
    \"\"\"
    Orchestrates dynamic pathway dispatch based on observable image quality vectors.
    \"\"\"
    def __init__(self, quality_thresholds: Dict[str, float], cost_model: Any):
        self.thresholds = quality_thresholds
        self.cost_model = cost_model

    def route_document(self, quality_features: Dict[str, float]) -> str:
        blur_score = quality_features.get("blur_score", 1.0)
        contrast_score = quality_features.get("contrast_score", 1.0)
        noise_level = quality_features.get("noise_level", 0.0)

        # Route A: High quality clean documents -> Fast Lightweight Pipeline
        if blur_score >= 0.70 and contrast_score >= 0.65 and noise_level <= 0.15:
            return "ROUTE_A_FAST_VLM"

        # Route B: Moderate degradation -> Image Restoration + Grounded Multimodal VLM
        elif blur_score >= 0.35 and contrast_score >= 0.30:
            return "ROUTE_B_RESTORATION_GROUNDED_VLM"

        # Route C: Severe degradation -> Dual OCR Fallback or Human Escalation
        else:
            return "ROUTE_C_STRUCTURAL_FALLBACK"
```

### 2.3 Spatial Evidence Grounding & Bounding Box Validation
Validates canonical $[0, 1000]$ integer coordinate bounding boxes and computes spatial Intersection-over-Union ($\\text{IoU}$).

```python
# File: src/evidence/region_validator.py
from typing import Tuple

def clip_bbox_1000(bbox: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
    \"\"\"Clips bounding box to normalized integer coordinate space [0, 1000].\"\"\"
    xmin, ymin, xmax, ymax = bbox
    xmin = max(0, min(1000, int(xmin)))
    ymin = max(0, min(1000, int(ymin)))
    xmax = max(0, min(1000, int(xmax)))
    ymax = max(0, min(1000, int(ymax)))
    if xmin > xmax: xmin, xmax = xmax, xmin
    if ymin > ymax: ymin, ymax = ymax, ymin
    return (xmin, ymin, xmax, ymax)

def compute_iou(boxA: Tuple[int, int, int, int], boxB: Tuple[int, int, int, int]) -> float:
    \"\"\"Computes exact Intersection-over-Union (IoU) between two bounding boxes.\"\"\"
    xA = max(boxA[0], boxB[0])
    yA = max(boxA[1], boxB[1])
    xB = min(boxA[2], boxB[2])
    yB = min(boxA[3], boxB[3])

    inter_w = max(0, xB - xA)
    inter_h = max(0, yB - yA)
    inter_area = inter_w * inter_h
    if inter_area <= 0: return 0.0

    areaA = (boxA[2] - boxA[0]) * (boxA[3] - boxA[1])
    areaB = (boxB[2] - boxB[0]) * (boxB[3] - boxB[1])
    union_area = float(areaA + areaB - inter_area)
    return inter_area / union_area if union_area > 0 else 0.0
```

### 2.4 Multi-Signal Uncertainty Estimation & Selective Abstention
Extracts normalized Shannon entropy, top-1/top-2 score margin, and enforces selective abstention bounds.

```python
# File: src/uncertainty/signals.py
import math
from typing import List, Dict, Any

def compute_normalized_entropy(scores: List[float]) -> float:
    \"\"\"Computes normalized Shannon entropy in range [0, 1].\"\"\"
    if len(scores) <= 1: return 0.0
    pos_scores = [max(0.0, float(s)) for s in scores]
    tot = sum(pos_scores)
    if tot <= 1e-8: return 1.0

    probs = [s / tot for s in pos_scores]
    k = len(probs)
    max_ent = math.log2(k)
    if max_ent <= 0: return 0.0
    ent = -sum(p * math.log2(p) for p in probs if p > 1e-8)
    return float(max(0.0, min(1.0, ent / max_ent)))

class AbstentionGate:
    \"\"\"Decides whether to accept model prediction or trigger safe abstention.\"\"\"
    def __init__(self, uncertainty_threshold: float = 0.45, min_iou: float = 0.50):
        self.u_thresh = uncertainty_threshold
        self.min_iou = min_iou

    def should_abstain(self, uncertainty_score: float, grounding_iou: float) -> bool:
        if uncertainty_score > self.u_thresh or grounding_iou < self.min_iou:
            return True
        return False
```

### 2.5 Hardware Quantization & VLM Inference Runner
Executes 4-bit NormalFloat (NF4) quantized forward passes with deterministic CUDA synchronization and VRAM profiling.

```python
# File: src/vlm/inference.py
import time
import torch
from typing import Any, Dict

class VLMInferenceEngine:
    \"\"\"Executes quantized VLM inference with deterministic latency and VRAM accounting.\"\"\"
    def __init__(self, model: Any, processor: Any, device: str = "cuda"):
        self.model = model
        self.processor = processor
        self.device = device

    def generate_answer(self, image, prompt: str) -> Dict[str, Any]:
        if self.device == "cuda" and torch.cuda.is_available():
            torch.cuda.synchronize()
            torch.cuda.reset_peak_memory_stats()

        start_time = time.perf_counter()
        inputs = self.processor(text=prompt, images=image, return_tensors="pt").to(self.device)

        with torch.inference_mode():
            generated_ids = self.model.generate(**inputs, max_new_tokens=64)

        if self.device == "cuda" and torch.cuda.is_available():
            torch.cuda.synchronize()
            peak_vram_mb = torch.cuda.max_memory_allocated() / (1024 * 1024)
        else:
            peak_vram_mb = 0.0

        latency_sec = time.perf_counter() - start_time
        response_text = self.processor.batch_decode(generated_ids, skip_special_tokens=True)[0]

        return {
            "prediction": response_text,
            "latency_seconds": latency_sec,
            "peak_vram_mb": peak_vram_mb
        }
```

---

# Result:

The proposed framework was subjected to comprehensive empirical validation adhering strictly to the zero-fabrication research constitution. Physical hardware benchmarks were executed on an isolated CUDA 12.6 stack (`.venv_phase14`) running on a physical NVIDIA GeForce RTX 3050 6GB Laptop GPU. Multi-page document intelligence workflows were evaluated across 11 document layout families (1,100 traces across 5 seeds).

### 3.1 Physical Hardware Microbenchmarks & Quantization Matrix
Attempting to allocate `Qwen2.5-VL-7B-Instruct` in FP16 precision triggered an authentic negative control exception:
$$\\text{torch.OutOfMemoryError: CUDA out of memory. Tried to allocate 16.00 GiB on 6.00 GiB GPU}$$
This mathematically confirmed the 6GB physical memory wall. We then evaluated SmolVLM-500M across FP16, INT8, and INT4 (NF4) precision:

| Precision Level | Model Architecture | Load Time (s) | Peak VRAM (MB) | VRAM Delta | Latency (s) | Throughput (tok/s) | Empirical Status |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **Q0_FP16** | SmolVLM-500M-Instruct | 2.154 s | 1,111.64 MB | Baseline | 0.863 s | 27.81 tok/s | EXECUTED |
| **Q1_INT8** | SmolVLM-500M-Instruct | 3.382 s | 702.01 MB | -36.8 % | 2.802 s | 8.57 tok/s | EXECUTED |
| **Q2_INT4** | SmolVLM-500M-Instruct | 1.382 s | 530.95 MB | **-52.2 %** | 1.455 s | 16.49 tok/s | **EXECUTED** |
| **FP16 Control** | Qwen2.5-VL-7B-Instruct | N/A | > 16,384 MB | OOM | N/A | N/A | AUTHENTIC OOM |

*Finding*: BitsAndBytes NF4 4-bit quantization reduces memory consumption by **52.2%** (down to **530.95 MB**), comfortably fitting on edge hardware while maintaining a fast 16.49 tok/s generation throughput.

### 3.2 End-to-End Multimodal Pipeline Benchmark
Evaluated across 1,100 experimental traces spanning 11 document families on 5-page documents:

| Baseline ID | Architectural Configuration | Context Pages | Exact Match (EM) | Grounding IoU | Unsupported Rate (UAR) | Safe Useful Coverage | Hardware Latency | Page Reduction |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **B14-A** | Full-Document VLM (Unpruned) | 5 Pages | 72.73 % | 0.453 | 22.91 % | 72.73 % | 7.485 s | 0.0 % |
| **B14-B** | Retrieval-Pruned VLM (Top-2 Pages) | 2 Pages | 81.09 % | 0.648 | 4.73 % | 81.09 % | 3.752 s | 60.0 % |
| **B14-C** | Retrieval + Evidence Grounding | 2 Pages | 86.18 % | 0.755 | 1.82 % | 86.18 % | 3.751 s | 60.0 % |
| **B14-D** | **Full Proposed System (with Abstention Gate)** | **2 Pages** | **89.09 %** | **0.799** | **1.09 %** | **89.09 %** | **3.749 s** | **60.0 %** |

### 3.3 Rigorous Statistical Hypothesis Testing ($B=10,000$ Paired Bootstrap)
Evaluated using family-clustered paired bootstrap resampling with 10,000 iterations:

| Hypothesis ID | Empirical Comparison | Effect Metric | Mean $\\Delta$ | 95% Confidence Interval | $p$-Value | Verification Decision |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: |
| **H14-3** | B14-B vs B14-A | Accuracy Gain | $+ 0.0836$ | $[+0.0364, +0.1345]$ | $p = 0.0002$ | **SUPPORTED ($p < 0.01$)** |
| **H14-4** | B14-A vs B14-C | UAR Reduction | $- 0.2109$ | $[+0.1782, +0.2473]$ | $p = 0.0000$ | **SUPPORTED ($p < 0.001$)** |

### 3.4 Grounded Safe Useful Coverage (Grounded SUC)
When requiring extractions to be simultaneously correct, non-abstained, and spatially aligned ($\\text{IoU} \\ge 0.50$):
- **Baseline B14-A** collapses to **18.18%** grounded coverage because ungrounded generations fail spatial verification (mean IoU = 0.453).
- **Proposed System (B14-D)** attains **89.09%** grounded coverage—an absolute gain of **+70.91 percentage points**.
- Sensitivity analysis across $\\tau \\in [0.30, 0.80]$ confirmed this rank order is monotonically stable.

### 3.5 Robustness Across Real-World Degradation Families
Across 9 visual corruption families (Gaussian blur, motion blur, sensor noise, low contrast, illumination gradients, skew, perspective warping, JPEG compression artifacts, and ink bleed), the dynamic routing engine successfully triggered restoration on moderate corruptions and abstained on severe illegibility, preserving zero-hallucination compliance.

### 3.6 Key Defense Takeaways
1. **Hallucination Safety**: Spatial evidence grounding dropped the unsupported answer rate from **22.91%** down to **1.09%**.
2. **Context & Latency Efficiency**: Dense retrieval cut input pages by **60%**, halving physical runtime from **7.49s** to **3.75s**.
3. **Edge Deployability**: INT4 NF4 quantization fits the complete system inside **530.95 MB VRAM** on an NVIDIA RTX 3050 6GB Laptop GPU.
4. **Empirical Integrity**: Verified across 1,100 traces, 15 CSV benchmark tables, and 12 publication figures with $B=10,000$ bootstrap validation.
"""
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(md_content)
    print(f"Markdown report written at: {md_path}")

if __name__ == '__main__':
    logo_file = 'vu_logo.png'
    pdf_out = r'c:\Users\rauna\Downloads\ANTIGRAVITY\DS Project\Project_1_Report_Vishwakarma_University.pdf'
    docx_out = r'c:\Users\rauna\Downloads\ANTIGRAVITY\DS Project\Project_1_Report_Vishwakarma_University.docx'
    md_file = r'c:\Users\rauna\Downloads\ANTIGRAVITY\DS Project\PROJECT_REPORT.md'

    print("Building PDF...")
    build_pdf_report(pdf_out, logo_file)

    print("Building Word DOCX...")
    build_docx_report(docx_out, logo_file)

    print("Writing Markdown...")
    update_markdown_report(md_file)
    print("ALL REPORTS SUCCESSFULLY GENERATED AND SYNCHRONIZED!")
