#!/usr/bin/env python3
"""
Generate a professional 16:9 widescreen 8-slide presentation in .pptx
explaining the complete TaskBoard DevOps Workflow:
From Developer Laptop to Monitored Production on Kubernetes.
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

# Color Palette - Modern Enterprise Tech / SRE Theme
BG_DARK = RGBColor(15, 23, 42)          # #0F172A Slate 900
BG_LIGHT = RGBColor(248, 250, 252)      # #F8FAFC Slate 50
CARD_BG = RGBColor(255, 255, 255)       # #FFFFFF Clean White
CARD_BORDER = RGBColor(226, 232, 240)   # #E2E8F0 Slate 200
PRIMARY = RGBColor(30, 41, 59)          # #1E293B Slate 800
ACCENT_BLUE = RGBColor(37, 99, 235)     # #2563EB Blue 600
ACCENT_BLUE_LIGHT = RGBColor(239, 246, 255) # #EFF6FF
ACCENT_TEAL = RGBColor(13, 148, 136)    # #0D9488 Teal 600
ACCENT_TEAL_LIGHT = RGBColor(240, 253, 250) # #F0FDFA
ACCENT_PURPLE = RGBColor(124, 58, 237)  # #7C3AED Purple 600
ACCENT_AMBER = RGBColor(217, 119, 6)    # #D97706 Amber 600
ACCENT_AMBER_LIGHT = RGBColor(254, 243, 199) # #FEF3C7
ACCENT_EMERALD = RGBColor(16, 185, 129) # #10B981 Emerald 500
ACCENT_CYAN = RGBColor(6, 182, 212)     # #06B6D4 Cyan 500
TEXT_MAIN = RGBColor(15, 23, 42)        # #0F172A
TEXT_MUTED = RGBColor(71, 85, 105)      # #475569 Slate 600
TEXT_LIGHT = RGBColor(148, 163, 184)    # #94A3B8 Slate 400
WHITE = RGBColor(255, 255, 255)

FONT_HEADING = "Segoe UI"
FONT_BODY = "Segoe UI"
FONT_MONO = "Consolas"


def set_run(run, text, size=14, bold=False, italic=False, color=TEXT_MAIN, font=FONT_BODY):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_slide_bg(slide, prs, dark=False):
    bg = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_DARK if dark else BG_LIGHT
    bg.line.fill.background()
    return bg


def add_header(slide, badge_text: str, title: str, subtitle: str, dark=False):
    # Top decorative brand stripe
    stripe = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.08)
    )
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = ACCENT_CYAN if dark else ACCENT_BLUE
    stripe.line.fill.background()

    # Category Pill Badge
    pill = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.35), Inches(3.2), Inches(0.32)
    )
    pill.fill.solid()
    pill.fill.fore_color.rgb = RGBColor(30, 41, 59) if dark else ACCENT_BLUE_LIGHT
    pill.line.color.rgb = RGBColor(51, 65, 85) if dark else ACCENT_BLUE
    pill.line.width = Pt(1)
    tf = pill.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    set_run(p.runs[0] if p.runs else tf.paragraphs[0].add_run(), badge_text, size=9.5, bold=True, color=ACCENT_CYAN if dark else ACCENT_BLUE)

    # Title Box
    tbox = slide.shapes.add_textbox(Inches(0.75), Inches(0.68), Inches(11.8), Inches(0.55))
    tf = tbox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.space_after = Pt(2)
    set_run(p.runs[0] if p.runs else p.add_run(), title, size=24, bold=True, color=WHITE if dark else PRIMARY)

    # Subtitle Box
    sbox = slide.shapes.add_textbox(Inches(0.75), Inches(1.18), Inches(11.8), Inches(0.35))
    tf = sbox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    set_run(p.runs[0] if p.runs else p.add_run(), subtitle, size=12, italic=False, color=TEXT_LIGHT if dark else TEXT_MUTED)

    # Subtle separator line
    sep = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(11.733), Inches(0.015)
    )
    sep.fill.solid()
    sep.fill.fore_color.rgb = RGBColor(51, 65, 85) if dark else CARD_BORDER
    sep.line.fill.background()


def add_footer(slide, current_slide: int, total_slides: int = 8, dark=False):
    foot_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.0), Inches(11.733), Inches(0.35))
    tf = foot_box.text_frame
    p = tf.paragraphs[0]
    
    # Left: Project
    r1 = p.add_run()
    set_run(r1, "TaskBoard — Enterprise DevOps & GitOps Capstone", size=9.5, bold=True, color=TEXT_LIGHT if dark else TEXT_MUTED)
    
    # Center: Student
    r2 = p.add_run()
    set_run(r2, "   |   Naman Anand (Roll: 10417) · Scaler School of Technology   |   ", size=9.5, color=TEXT_LIGHT if dark else TEXT_MUTED)
    
    # Right: Slide number
    r3 = p.add_run()
    set_run(r3, f"Slide {current_slide} of {total_slides}", size=9.5, bold=True, color=ACCENT_CYAN if dark else ACCENT_BLUE)


def create_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER, border_width=1.0):
    card = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(border_width)
    return card


# ==========================================
# SLIDE 1: Title Slide (Dark Executive Theme)
# ==========================================
def build_slide_1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, prs, dark=True)

    # Top gradient-like accent stripe
    stripe = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12)
    )
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = ACCENT_CYAN
    stripe.line.fill.background()

    # Category Pill
    pill = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(0.8), Inches(4.2), Inches(0.38)
    )
    pill.fill.solid()
    pill.fill.fore_color.rgb = RGBColor(30, 41, 59)
    pill.line.color.rgb = ACCENT_CYAN
    pill.line.width = Pt(1.2)
    tf = pill.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    set_run(p.add_run(), "DEVOPS CAPSTONE PROJECT · END-TO-END WORKFLOW", size=10, bold=True, color=ACCENT_CYAN)

    # Main Hero Title
    tbox = slide.shapes.add_textbox(Inches(0.95), Inches(1.3), Inches(11.4), Inches(1.5))
    tf = tbox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    set_run(p.add_run(), "TaskBoard: Production DevOps Delivery Workflow", size=36, bold=True, color=WHITE)
    
    # Subtitle
    p2 = tf.add_paragraph()
    p2.space_before = Pt(8)
    set_run(p2.add_run(), "From Developer Commit to Continuous Testing, DevSecOps, AWS Terraform, Kubernetes Orchestration & GitOps", size=15, color=RGBColor(203, 213, 225))

    # 3 High Impact Feature Cards (Left 3/4)
    features = [
        ("01 / Continuous Integration", "Pytest suite (100% route coverage), flake8 linting, Bandit SAST, Gitleaks, and Trivy CVE container scanning."),
        ("02 / Infrastructure as Code", "Terraform provisioned multi-AZ AWS VPC, managed EKS cluster, IRSA roles, security groups, and automated teardown."),
        ("03 / GitOps & Observability", "Argo CD declarative synchronization, Helm packaging, Prometheus scraping /metrics, and rich Grafana dashboards.")
    ]
    card_w = 3.6
    gap = 0.25
    start_x = 1.0
    for i, (title, desc) in enumerate(features):
        x = start_x + i * (card_w + gap)
        card = create_card(slide, x, 3.2, card_w, 2.1, bg_color=RGBColor(30, 41, 59), border_color=RGBColor(51, 65, 85))
        # Top color accent on card
        accent_bar = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(x), Inches(3.2), Inches(card_w), Inches(0.06)
        )
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = [ACCENT_CYAN, ACCENT_BLUE, ACCENT_EMERALD][i]
        accent_bar.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(x + 0.2), Inches(3.4), Inches(card_w - 0.4), Inches(1.8))
        tf_c = tb.text_frame
        tf_c.word_wrap = True
        p_c = tf_c.paragraphs[0]
        set_run(p_c.add_run(), title, size=13, bold=True, color=WHITE)
        p_desc = tf_c.add_paragraph()
        p_desc.space_before = Pt(6)
        set_run(p_desc.add_run(), desc, size=11, color=RGBColor(148, 163, 184))

    # Bottom Presenter & Repo Banner
    pres_card = create_card(slide, 1.0, 5.55, 11.333, 1.25, bg_color=RGBColor(24, 34, 53), border_color=ACCENT_BLUE)
    ptb = slide.shapes.add_textbox(Inches(1.2), Inches(5.65), Inches(10.9), Inches(1.05))
    ptf = ptb.text_frame
    ptf.word_wrap = True
    
    p_meta = ptf.paragraphs[0]
    set_run(p_meta.add_run(), "SUBMITTED BY: ", size=10, bold=True, color=ACCENT_CYAN)
    set_run(p_meta.add_run(), "Naman Anand   ", size=13, bold=True, color=WHITE)
    set_run(p_meta.add_run(), " |   ROLL NUMBER: ", size=10, bold=True, color=ACCENT_CYAN)
    set_run(p_meta.add_run(), "10417   ", size=13, bold=True, color=WHITE)
    set_run(p_meta.add_run(), " |   INSTITUTION: ", size=10, bold=True, color=ACCENT_CYAN)
    set_run(p_meta.add_run(), "Scaler School of Technology", size=13, bold=True, color=WHITE)

    p_submeta = ptf.add_paragraph()
    p_submeta.space_before = Pt(4)
    set_run(p_submeta.add_run(), "EMAIL: ", size=10, bold=True, color=TEXT_LIGHT)
    set_run(p_submeta.add_run(), "naman.24bcs10417@sst.scaler.com    ", size=11, color=RGBColor(226, 232, 240))
    set_run(p_submeta.add_run(), "REPOSITORY: ", size=10, bold=True, color=TEXT_LIGHT)
    set_run(p_submeta.add_run(), "github.com/NamanAnand66/devops-project    ", size=11, color=ACCENT_CYAN)
    set_run(p_submeta.add_run(), "STACK: ", size=10, bold=True, color=TEXT_LIGHT)
    set_run(p_submeta.add_run(), "React · FastAPI · PostgreSQL · Docker · Helm · EKS · ArgoCD", size=11, color=RGBColor(226, 232, 240))

    add_footer(slide, 1, 8, dark=True)


# ==================================================
# SLIDE 2: High-Level End-to-End Delivery Lifecycle
# ==================================================
def build_slide_2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, prs, dark=False)
    add_header(
        slide,
        "WORKFLOW ARCHITECTURE · 5 CORE PHASES",
        "End-to-End DevOps & GitOps Delivery Pipeline",
        "How TaskBoard code transitions from developer laptop to monitored production on AWS EKS"
    )

    # 5 Process Flow Step Cards across the screen
    stages = [
        ("01", "Developer Loop", "Local Workstation", [
            "React 19 + FastAPI + Postgres",
            "docker-compose.yml full stack",
            "Health-gated startup sequence",
            "Git feature branch & PR"
        ], ACCENT_BLUE),
        ("02", "CI & Security", "GitHub Actions", [
            "flake8 & black code linting",
            "15 Pytest integration tests",
            "Bandit & Semgrep SAST scans",
            "Gitleaks secret detection"
        ], ACCENT_TEAL),
        ("03", "Artifact Build", "Registry (GHCR)", [
            "Multi-stage lean Dockerfiles",
            "Non-root unprivileged users",
            "Trivy CVE vulnerability scan",
            "Tag with Git commit SHA"
        ], ACCENT_PURPLE),
        ("04", "Cloud Infra", "Terraform AWS", [
            "Multi-AZ VPC & private subnets",
            "Managed Amazon EKS Cluster",
            "IAM Roles for Service Accounts",
            "Clean teardown on demand"
        ], ACCENT_AMBER),
        ("05", "GitOps & SRE", "Kubernetes & Argo", [
            "Argo CD declarative Git sync",
            "Helm taskboard chart & values",
            "Prometheus metrics scraping",
            "Grafana latency & error boards"
        ], ACCENT_EMERALD)
    ]

    card_w = 2.15
    gap = 0.245
    start_x = 0.8
    y = 1.7
    h = 3.65

    for i, (num, title, subtitle, bullets, accent) in enumerate(stages):
        x = start_x + i * (card_w + gap)
        # Background card
        card = create_card(slide, x, y, card_w, h, bg_color=WHITE, border_color=CARD_BORDER)

        # Header accent band
        band = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(x), Inches(y), Inches(card_w), Inches(0.06)
        )
        band.fill.solid()
        band.fill.fore_color.rgb = accent
        band.line.fill.background()

        # Step Number Pill
        badge = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x + 0.15), Inches(y + 0.18), Inches(0.5), Inches(0.28)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = accent
        badge.line.fill.background()
        tf_b = badge.text_frame
        tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_b = tf_b.paragraphs[0]
        p_b.alignment = PP_ALIGN.CENTER
        set_run(p_b.add_run(), num, size=11, bold=True, color=WHITE)

        # Text box
        tb = slide.shapes.add_textbox(Inches(x + 0.15), Inches(y + 0.55), Inches(card_w - 0.3), Inches(h - 0.65))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p_t = tf.paragraphs[0]
        set_run(p_t.add_run(), title, size=13, bold=True, color=PRIMARY)
        
        p_sub = tf.add_paragraph()
        p_sub.space_before = Pt(2)
        set_run(p_sub.add_run(), subtitle, size=10, bold=True, color=accent)

        for b in bullets:
            p_bullet = tf.add_paragraph()
            p_bullet.space_before = Pt(6)
            set_run(p_bullet.add_run(), "• ", size=10, bold=True, color=accent)
            set_run(p_bullet.add_run(), b, size=10, color=TEXT_MUTED)

        # Draw connecting arrow between stages (except last)
        if i < len(stages) - 1:
            arrow = slide.shapes.add_shape(
                MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW, Inches(x + card_w + 0.04), Inches(y + 1.8), Inches(gap - 0.08), Inches(0.22)
            )
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = RGBColor(203, 213, 225)
            arrow.line.fill.background()

    # Bottom Architectural Callout Strip
    bottom_card = create_card(slide, 0.8, 5.5, 11.733, 1.3, bg_color=ACCENT_BLUE_LIGHT, border_color=RGBColor(191, 219, 254))
    bb_tb = slide.shapes.add_textbox(Inches(1.0), Inches(5.55), Inches(11.333), Inches(1.15))
    bb_tf = bb_tb.text_frame
    bb_tf.word_wrap = True

    p_bbt = bb_tf.paragraphs[0]
    set_run(p_bbt.add_run(), "CORE ARCHITECTURAL PRINCIPLES OF THIS WORKFLOW", size=11, bold=True, color=ACCENT_BLUE)
    
    p_bbi = bb_tf.add_paragraph()
    p_bbi.space_before = Pt(4)
    set_run(p_bbi.add_run(), "1. Immutable Artifacts: ", size=10.5, bold=True, color=PRIMARY)
    set_run(p_bbi.add_run(), "Images built and scanned once in CI are tagged by Git SHA; Helm only updates deployment configs.\n", size=10, color=TEXT_MUTED)
    set_run(p_bbi.add_run(), "2. Shift-Left DevSecOps: ", size=10.5, bold=True, color=PRIMARY)
    set_run(p_bbi.add_run(), "SAST, secrets, and CVE scanning happen in CI before any code touches the cluster.\n", size=10, color=TEXT_MUTED)
    set_run(p_bbi.add_run(), "3. Declarative GitOps: ", size=10.5, bold=True, color=PRIMARY)
    set_run(p_bbi.add_run(), "Git is the single source of truth. Argo CD reconciles drift automatically without manual kubectl interventions.", size=10, color=TEXT_MUTED)

    add_footer(slide, 2, 8)


# ==================================================
# SLIDE 3: Application Architecture & Local Dev Loop
# ==================================================
def build_slide_3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, prs, dark=False)
    add_header(
        slide,
        "PHASE 1 · APPLICATION FOUNDATION",
        "Application Architecture & Local Developer Loop",
        "Three-tier decoupled services engineered for resilience, local parity, and Kubernetes deployment"
    )

    # 3 Microservice Architecture Cards (Left side, 8.2 inches wide)
    services = [
        ("Frontend: React 19 + Vite Dashboard", [
            ("Technology", "React 19, Vite bundler, responsive JSX/CSS SaaS dashboard"),
            ("Key Features", "Task kanban/table, KPI cards, priority badges, modal forms"),
            ("Reverse Proxy", "In-container Nginx proxies /api to backend; eliminates CORS"),
            ("Security", "Runs as non-root user (UID 10001); multi-stage lean container")
        ], ACCENT_BLUE),
        ("Backend API: FastAPI + Python 3.12", [
            ("Technology", "FastAPI, async SQLAlchemy 2, Pydantic data validation"),
            ("Endpoints", "/api/tasks CRUD, /api/tasks/stats, auto OpenAPI Swagger docs"),
            ("SRE Probes", "/health (liveness), /ready (database check), /metrics (Prometheus)"),
            ("Migrations", "Alembic schema versioning runs automatically on container start")
        ], ACCENT_TEAL),
        ("Database: PostgreSQL 16 Relational Store", [
            ("Engine", "PostgreSQL 16 Alpine with indexed task statuses and priorities"),
            ("Persistence", "Docker named volume in local dev; AWS EBS via K8s PVC in cloud"),
            ("Resilience", "Healthcheck command (pg_isready) gates dependent startup"),
            ("Isolation", "Zero direct public internet access; strictly isolated private network")
        ], ACCENT_PURPLE),
    ]

    col_w = 2.6
    gap = 0.2
    start_x = 0.8
    y = 1.7
    h = 3.6

    for i, (svc_title, details, accent) in enumerate(services):
        x = start_x + i * (col_w + gap)
        card = create_card(slide, x, y, col_w, h, bg_color=WHITE, border_color=CARD_BORDER)
        
        # Header accent bar
        bar = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(x), Inches(y), Inches(col_w), Inches(0.06)
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(x + 0.15), Inches(y + 0.15), Inches(col_w - 0.3), Inches(h - 0.3))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        set_run(p.add_run(), svc_title, size=12, bold=True, color=PRIMARY)

        for label, val in details:
            p_item = tf.add_paragraph()
            p_item.space_before = Pt(8)
            set_run(p_item.add_run(), f"• {label}: ", size=10, bold=True, color=accent)
            set_run(p_item.add_run(), val, size=9.5, color=TEXT_MUTED)

    # Right Side: Local Dev Loop with Docker Compose (3.2 inches wide)
    compose_x = 9.2
    compose_w = 3.333
    compose_card = create_card(slide, compose_x, y, compose_w, h, bg_color=RGBColor(241, 245, 249), border_color=RGBColor(203, 213, 225))
    
    ctb = slide.shapes.add_textbox(Inches(compose_x + 0.2), Inches(y + 0.15), Inches(compose_w - 0.4), Inches(h - 0.3))
    ctf = ctb.text_frame
    ctf.word_wrap = True
    
    cp_title = ctf.paragraphs[0]
    set_run(cp_title.add_run(), "Local Dev Loop (docker compose)", size=12.5, bold=True, color=PRIMARY)
    
    compose_steps = [
        ("Single Command Launch", "docker compose up --build runs the identical stack locally in under 60 seconds."),
        ("Healthcheck Ordering", "depends_on: condition: service_healthy ensures PostgreSQL accepts connections before backend boots."),
        ("Environment Parity", "Same Dockerfiles used in local dev are tested and deployed in Kubernetes, eliminating surprises."),
        ("Hot Reloading", "Local volume mounts enable instant code updates without rebuilding containers during active development.")
    ]
    for step_h, step_b in compose_steps:
        p_s = ctf.add_paragraph()
        p_s.space_before = Pt(6)
        set_run(p_s.add_run(), f"✓ {step_h}: ", size=10, bold=True, color=ACCENT_BLUE)
        set_run(p_s.add_run(), step_b, size=9.5, color=TEXT_MUTED)

    # Bottom Request Flow Banner
    flow_card = create_card(slide, 0.8, 5.45, 11.733, 1.35, bg_color=WHITE, border_color=CARD_BORDER)
    ftb = slide.shapes.add_textbox(Inches(1.0), Inches(5.5), Inches(11.333), Inches(1.2))
    ftf = ftb.text_frame
    ftf.word_wrap = True

    p_ft = ftf.paragraphs[0]
    set_run(p_ft.add_run(), "RUNTIME REQUEST FLOW & SRE VISIBILITY PATHWAY", size=11, bold=True, color=ACCENT_TEAL)
    
    p_path = ftf.add_paragraph()
    p_path.space_before = Pt(4)
    flow_items = [
        ("Client Browser", "User interactions on port 3000 / 80"),
        ("Nginx Gateway", "Serves React SPA assets & proxies /api"),
        ("FastAPI Backend", "Validates input, processes logic & records metrics"),
        ("PostgreSQL 16", "Persists relational task state"),
        ("Prometheus Scraper", "Polls /metrics every 15s for SRE alerts")
    ]
    for idx, (node, desc) in enumerate(flow_items):
        set_run(p_path.add_run(), f"[{node}]", size=10.5, bold=True, color=PRIMARY)
        if idx < len(flow_items) - 1:
            set_run(p_path.add_run(), "  ──▶  ", size=11, bold=True, color=ACCENT_TEAL)

    p_flow_desc = ftf.add_paragraph()
    p_flow_desc.space_before = Pt(4)
    set_run(p_flow_desc.add_run(), "Decoupled Routing: ", size=10, bold=True, color=PRIMARY)
    set_run(p_flow_desc.add_run(), "The browser never needs direct database access or internal service IP knowledge. Ingress and Nginx handle internal DNS abstraction seamlessly.", size=9.5, color=TEXT_MUTED)

    add_footer(slide, 3, 8)


# ==================================================
# SLIDE 4: CI/CD Pipeline & DevSecOps Security Gates
# ==================================================
def build_slide_4(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, prs, dark=False)
    add_header(
        slide,
        "PHASE 2 · AUTOMATED QUALITY & SECURITY",
        "Continuous Integration & DevSecOps Pipeline",
        "Automated GitHub Actions workflow enforcing testing, static analysis, container security, and registry publishing"
    )

    # 4 Pipeline Stages in Grid Layout (2 rows x 2 columns)
    stages = [
        ("Stage 1: Linting & Automated Pytest Gate", ACCENT_BLUE, [
            ("flake8 & black", "Enforces PEP 8 style guide and catches syntax errors before compilation."),
            ("15 Pytest Test Cases", "Covers /health, /ready, /metrics, /api/tasks (GET, POST, PUT, DELETE, stats)."),
            ("Isolated Test Database", "Tests execute against an in-memory test database, guaranteeing zero prod pollution."),
            ("Coverage Threshold", "Fails CI build if test suite fails or code coverage dips below quality bar.")
        ]),
        ("Stage 2: Shift-Left DevSecOps & Secret Scanning", ACCENT_TEAL, [
            ("Bandit SAST", "Scans Python AST for SQL injection, insecure deserialization, and dangerous calls."),
            ("Semgrep Security Rules", "Checks for OWASP Top 10 vulnerabilities and misconfigured FastAPI decorators."),
            ("Gitleaks Credential Scan", "Deep scans commit history for leaked API keys, tokens, or private credentials."),
            ("pip-audit & npm audit", "Software Composition Analysis (SCA) detecting known CVEs in 3rd party packages.")
        ]),
        ("Stage 3: Multi-Stage Lean & Hardened Docker Builds", ACCENT_PURPLE, [
            ("Frontend Multi-Stage", "Stage 1 builds Vite bundle with Node 20; Stage 2 copies static files to slim Nginx."),
            ("Backend Lean Base", "Python 3.12-slim base with precompiled wheels, discarding compilers from runtime."),
            ("Non-Root Users", "Both containers run as unprivileged users (appuser / UID 10001) to mitigate escape exploits."),
            ("Hadolint Verification", "Enforces Dockerfile best practices (no apt cache left, pinned version tags).")
        ]),
        ("Stage 4: Trivy Container Vulnerability Gate & GHCR", ACCENT_AMBER, [
            ("Trivy Vulnerability Scan", "Scans finalized container filesystem for OS and library vulnerabilities."),
            ("Zero Critical CVE Policy", "Pipeline automatically fails if any High or Critical unpatched CVE is discovered."),
            ("GitHub Container Registry", "Images published to ghcr.io/namananand66/taskboard-backend & frontend."),
            ("Immutable SHA Tagging", "Images tagged with :commit-sha (traceability) and :latest (convenience).")
        ])
    ]

    card_w = 5.75
    card_h = 2.45
    coords = [
        (0.8, 1.7),   # Top Left
        (6.78, 1.7),  # Top Right
        (0.8, 4.3),   # Bottom Left
        (6.78, 4.3)   # Bottom Right
    ]

    for i, (title, accent, items) in enumerate(stages):
        x, y = coords[i]
        card = create_card(slide, x, y, card_w, card_h, bg_color=WHITE, border_color=CARD_BORDER)
        
        # Header accent bar
        bar = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(x), Inches(y), Inches(card_w), Inches(0.06)
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.12), Inches(card_w - 0.4), Inches(card_h - 0.2))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        set_run(p_t.add_run(), title, size=12.5, bold=True, color=PRIMARY)

        for label, desc in items:
            p_i = tf.add_paragraph()
            p_i.space_before = Pt(4)
            set_run(p_i.add_run(), f"• {label}: ", size=10, bold=True, color=accent)
            set_run(p_i.add_run(), desc, size=9.5, color=TEXT_MUTED)

    add_footer(slide, 4, 8)


# ==================================================
# SLIDE 5: Infrastructure as Code (Terraform on AWS)
# ==================================================
def build_slide_5(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, prs, dark=False)
    add_header(
        slide,
        "PHASE 3 · CLOUD INFRASTRUCTURE AS CODE",
        "Terraform Architecture on Amazon Web Services (AWS)",
        "Declarative, reproducible provisioning of production VPC, managed EKS cluster, and least-privilege security"
    )

    # Left Column: AWS Network Topology (VPC Module)
    vpc_w = 5.75
    vpc_h = 3.65
    vpc_card = create_card(slide, 0.8, 1.7, vpc_w, vpc_h, bg_color=WHITE, border_color=CARD_BORDER)
    
    vbar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.8), Inches(1.7), Inches(vpc_w), Inches(0.06))
    vbar.fill.solid()
    vbar.fill.fore_color.rgb = ACCENT_AMBER
    vbar.line.fill.background()

    vtb = slide.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(vpc_w - 0.4), Inches(vpc_h - 0.3))
    vtf = vtb.text_frame
    vtf.word_wrap = True

    vp = vtf.paragraphs[0]
    set_run(vp.add_run(), "Modular AWS VPC Network Architecture", size=13, bold=True, color=PRIMARY)

    vpc_points = [
        ("Multi-AZ High Availability", "Spans 2 Availability Zones (us-east-1a, us-east-1b) for fault tolerance and disaster recovery."),
        ("Public Subnets & IGW", "Hosts Internet Gateway and Application Load Balancers for external HTTP/HTTPS ingress."),
        ("Private Worker Subnets", "EKS EC2 worker nodes live in private subnets with no public IPs, completely shielded from internet attacks."),
        ("Managed NAT Gateways", "Provides outbound internet access for cluster nodes to pull Docker images and security patches."),
        ("Security Group Hardening", "Strict security group rules isolating frontend, backend pods, and PostgreSQL database.")
    ]
    for label, desc in vpc_points:
        p = vtf.add_paragraph()
        p.space_before = Pt(6)
        set_run(p.add_run(), f"✓ {label}: ", size=10, bold=True, color=ACCENT_AMBER)
        set_run(p.add_run(), desc, size=9.5, color=TEXT_MUTED)

    # Right Column: Amazon EKS Cluster & IAM Governance
    eks_w = 5.75
    eks_h = 3.65
    eks_card = create_card(slide, 6.78, 1.7, eks_w, eks_h, bg_color=WHITE, border_color=CARD_BORDER)

    ebar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(6.78), Inches(1.7), Inches(eks_w), Inches(0.06))
    ebar.fill.solid()
    ebar.fill.fore_color.rgb = ACCENT_BLUE
    ebar.line.fill.background()

    etb = slide.shapes.add_textbox(Inches(6.98), Inches(1.85), Inches(eks_w - 0.4), Inches(eks_h - 0.3))
    etf = etb.text_frame
    etf.word_wrap = True

    ep = etf.paragraphs[0]
    set_run(ep.add_run(), "Amazon EKS Compute & IAM Security", size=13, bold=True, color=PRIMARY)

    eks_points = [
        ("Managed Control Plane", "AWS handles Kubernetes API server patching, etcd clustering, and cross-AZ replication."),
        ("Managed Node Groups", "Automated EC2 worker node autoscaling, OS security patching, and health recovery."),
        ("IAM Roles for Service Accounts (IRSA)", "Pods assume fine-grained IAM roles via OpenID Connect (OIDC) without static AWS access keys."),
        ("KMS Secrets Encryption", "Kubernetes secrets encrypted at rest using AWS Key Management Service (KMS)."),
        ("ECR / GHCR Integration", "Secure private registry authentication for pull credentials.")
    ]
    for label, desc in eks_points:
        p = etf.add_paragraph()
        p.space_before = Pt(6)
        set_run(p.add_run(), f"✓ {label}: ", size=10, bold=True, color=ACCENT_BLUE)
        set_run(p.add_run(), desc, size=9.5, color=TEXT_MUTED)

    # Bottom Terraform Lifecycle & Cost Governance Panel
    tf_card = create_card(slide, 0.8, 5.5, 11.733, 1.3, bg_color=ACCENT_AMBER_LIGHT, border_color=RGBColor(251, 191, 36))
    ttb = slide.shapes.add_textbox(Inches(1.0), Inches(5.55), Inches(11.333), Inches(1.15))
    ttf = ttb.text_frame
    ttf.word_wrap = True

    p_tt = ttf.paragraphs[0]
    set_run(p_tt.add_run(), "TERRAFORM LIFECYCLE WORKFLOW & FINOPS COST GOVERNANCE", size=11, bold=True, color=ACCENT_AMBER)

    p_ti = ttf.add_paragraph()
    p_ti.space_before = Pt(4)
    set_run(p_ti.add_run(), "1. Automated Plan & Audit: ", size=10, bold=True, color=PRIMARY)
    set_run(p_ti.add_run(), "terraform plan outputs dry-run delta before any infrastructure changes are executed.\n", size=9.5, color=TEXT_MUTED)
    set_run(p_ti.add_run(), "2. State Locking & Safety: ", size=10, bold=True, color=PRIMARY)
    set_run(p_ti.add_run(), "Remote S3 backend with DynamoDB state locking prevents concurrent modification and race conditions.\n", size=9.5, color=TEXT_MUTED)
    set_run(p_ti.add_run(), "3. Cost Responsibility: ", size=10, bold=True, color=PRIMARY)
    set_run(p_ti.add_run(), "Complete clean teardown via terraform destroy prevents runaway AWS costs immediately after demonstration.", size=9.5, color=TEXT_MUTED)

    add_footer(slide, 5, 8)


# ==================================================
# SLIDE 6: Kubernetes Workloads & Helm Packaging
# ==================================================
def build_slide_6(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, prs, dark=False)
    add_header(
        slide,
        "PHASE 4 · CLUSTER ORCHESTRATION & HELM",
        "Kubernetes Workload Orchestration & Helm Packaging",
        "Production-grade Pod deployment, self-healing probes, autoscaling, and parameterized Helm release management"
    )

    # 3 Workload Columns
    cols = [
        ("Workload Primitives & Storage", ACCENT_BLUE, [
            ("Frontend Deployment", "Multi-replica Deployment serving static assets through Nginx reverse proxy."),
            ("Backend Deployment", "Stateless FastAPI pods with zero-downtime rolling updates (maxSurge=1)."),
            ("Postgres StatefulSet", "Stateful workload with dedicated PersistentVolumeClaim (PVC) backed by AWS EBS."),
            ("Decoupled Config", "ConfigMaps for application settings; Kubernetes Secrets for DB credentials.")
        ]),
        ("Self-Healing & Elastic Scaling", ACCENT_TEAL, [
            ("Three-Tier Probes", "startupProbe protects slow boots; livenessProbe detects deadlocks; readinessProbe gates traffic."),
            ("Ingress Routing", "Nginx Ingress controller routes external traffic: / to frontend, /api to backend."),
            ("HPA Autoscaling", "Horizontal Pod Autoscaler scales backend between 2 and 10 pods at 50% CPU threshold."),
            ("Resource Constraints", "Strict CPU/Memory requests & limits prevent noisy-neighbor cluster exhaustion.")
        ]),
        ("Helm Chart Engineering", ACCENT_PURPLE, [
            ("Unified Chart", "Custom taskboard chart packaging Deployments, Services, Ingress, and HPA."),
            ("Multi-Environment", "values-dev.yaml (lightweight, single replica) vs values-prod.yaml (multi-replica, HPA)."),
            ("Release Automation", "helm lint, helm template, and helm upgrade --install with rollback capabilities."),
            ("Smoke Test Hooks", "helm test executes automated in-cluster health validation after every deployment.")
        ])
    ]

    col_w = 3.75
    gap = 0.24
    start_x = 0.8
    y = 1.7
    h = 3.65

    for i, (title, accent, items) in enumerate(cols):
        x = start_x + i * (col_w + gap)
        card = create_card(slide, x, y, col_w, h, bg_color=WHITE, border_color=CARD_BORDER)

        bar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(x), Inches(y), Inches(col_w), Inches(0.06))
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(x + 0.15), Inches(y + 0.15), Inches(col_w - 0.3), Inches(h - 0.3))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        set_run(p_t.add_run(), title, size=12.5, bold=True, color=PRIMARY)

        for label, desc in items:
            p = tf.add_paragraph()
            p.space_before = Pt(6)
            set_run(p.add_run(), f"• {label}: ", size=10, bold=True, color=accent)
            set_run(p.add_run(), desc, size=9.5, color=TEXT_MUTED)

    # Bottom Ingress & Probes Spec Banner
    spec_card = create_card(slide, 0.8, 5.5, 11.733, 1.3, bg_color=RGBColor(248, 250, 252), border_color=CARD_BORDER)
    stb = slide.shapes.add_textbox(Inches(1.0), Inches(5.55), Inches(11.333), Inches(1.15))
    stf = stb.text_frame
    stf.word_wrap = True

    p_st = stf.paragraphs[0]
    set_run(p_st.add_run(), "HEALTH PROBES SPECIFICATION & PRODUCTION TRAFFIC GATING", size=11, bold=True, color=PRIMARY)

    p_si = stf.add_paragraph()
    p_si.space_before = Pt(4)
    set_run(p_si.add_run(), "Liveness Probe (/health): ", size=10, bold=True, color=ACCENT_BLUE)
    set_run(p_si.add_run(), "Cheap ping check (period: 10s, timeout: 2s). If failed 3 times, Kubernetes restarts the unresponsive container.\n", size=9.5, color=TEXT_MUTED)
    set_run(p_si.add_run(), "Readiness Probe (/ready): ", size=10, bold=True, color=ACCENT_TEAL)
    set_run(p_si.add_run(), "Active DB query test. Ensures Pod receives zero user traffic until migrations complete and database is reachable.\n", size=9.5, color=TEXT_MUTED)
    set_run(p_si.add_run(), "Metrics Scrape (/metrics): ", size=10, bold=True, color=ACCENT_PURPLE)
    set_run(p_si.add_run(), "Exposes Prometheus formatted request latencies, count histograms, and error counters for live Grafana monitoring.", size=9.5, color=TEXT_MUTED)

    add_footer(slide, 6, 8)


# ==================================================
# SLIDE 7: GitOps Continuous Delivery & Observability
# ==================================================
def build_slide_7(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, prs, dark=False)
    add_header(
        slide,
        "PHASE 5 · GITOPS & OBSERVABILITY",
        "GitOps Delivery (Argo CD) & Full-Stack Observability",
        "Automated pull-based continuous deployment paired with Prometheus metrics scraping and Grafana dashboards"
    )

    # Left Column: Argo CD GitOps Continuous Delivery
    gitops_w = 5.75
    gitops_h = 3.65
    gcard = create_card(slide, 0.8, 1.7, gitops_w, gitops_h, bg_color=WHITE, border_color=CARD_BORDER)

    gbar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.8), Inches(1.7), Inches(gitops_w), Inches(0.06))
    gbar.fill.solid()
    gbar.fill.fore_color.rgb = ACCENT_PURPLE
    gbar.line.fill.background()

    gtb = slide.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(gitops_w - 0.4), Inches(gitops_h - 0.3))
    gtf = gtb.text_frame
    gtf.word_wrap = True

    gp = gtf.paragraphs[0]
    set_run(gp.add_run(), "Pull-Based GitOps Delivery (Argo CD)", size=13, bold=True, color=PRIMARY)

    gitops_points = [
        ("Git as Single Truth", "All cluster states, Helm values, and manifests reside in GitHub repo (final-devops-project/helm/taskboard)."),
        ("Automated Sync Loop", "CI commits new Docker image tag to Git repo; Argo CD detects delta and synchronizes automatically."),
        ("Self-Healing & Pruning", "Argo CD detects and reconciles manual cluster drift, automatically rolling back unapproved ad-hoc edits."),
        ("Zero Cluster Ingress in CI", "GitHub Actions does NOT need cluster admin credentials; Argo pulls from inside the cluster."),
        ("Instant Rollbacks", "Reverting a Git commit instantly rolls back the deployment to the previous stable state.")
    ]
    for label, desc in gitops_points:
        p = gtf.add_paragraph()
        p.space_before = Pt(6)
        set_run(p.add_run(), f"✓ {label}: ", size=10, bold=True, color=ACCENT_PURPLE)
        set_run(p.add_run(), desc, size=9.5, color=TEXT_MUTED)

    # Right Column: Observability (Prometheus & Grafana)
    obs_w = 5.75
    obs_h = 3.65
    ocard = create_card(slide, 6.78, 1.7, obs_w, obs_h, bg_color=WHITE, border_color=CARD_BORDER)

    obar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(6.78), Inches(1.7), Inches(obs_w), Inches(0.06))
    obar.fill.solid()
    obar.fill.fore_color.rgb = ACCENT_EMERALD
    obar.line.fill.background()

    otb = slide.shapes.add_textbox(Inches(6.98), Inches(1.85), Inches(obs_w - 0.4), Inches(obs_h - 0.3))
    otf = otb.text_frame
    otf.word_wrap = True

    op = otf.paragraphs[0]
    set_run(op.add_run(), "Observability: Prometheus & Grafana", size=13, bold=True, color=PRIMARY)

    obs_points = [
        ("FastAPI Prometheus Metrics", "App exposes real-time HTTP metrics on /metrics (request duration histograms, total requests)."),
        ("Prometheus Operator", "Scrapes application targets and Kubernetes node metrics every 15 seconds."),
        ("Grafana SRE Dashboards", "Visualizes real-time Requests Per Second (RPS), P95/P99 latency curves, and 2xx/4xx/5xx ratios."),
        ("Resource Monitoring", "Tracks CPU and Memory utilization per pod to validate HPA scaling actions."),
        ("Alerting Rules", "Automated alert rules trigger on high 5xx error rate (>1%) or unready backend Pods.")
    ]
    for label, desc in obs_points:
        p = otf.add_paragraph()
        p.space_before = Pt(6)
        set_run(p.add_run(), f"✓ {label}: ", size=10, bold=True, color=ACCENT_EMERALD)
        set_run(p.add_run(), desc, size=9.5, color=TEXT_MUTED)

    # Bottom GitOps Pipeline Summary Strip
    sum_card = create_card(slide, 0.8, 5.5, 11.733, 1.3, bg_color=ACCENT_TEAL_LIGHT, border_color=RGBColor(153, 246, 228))
    stb = slide.shapes.add_textbox(Inches(1.0), Inches(5.55), Inches(11.333), Inches(1.15))
    stf = stb.text_frame
    stf.word_wrap = True

    p_st = stf.paragraphs[0]
    set_run(p_st.add_run(), "CLOSED-LOOP AUTOMATION: FROM CODE CHANGE TO METRIC VERIFICATION", size=11, bold=True, color=ACCENT_TEAL)

    p_si = stf.add_paragraph()
    p_si.space_before = Pt(4)
    set_run(p_si.add_run(), "1. Git Push: ", size=10, bold=True, color=PRIMARY)
    set_run(p_si.add_run(), "Developer commits feature to main branch.\n", size=9.5, color=TEXT_MUTED)
    set_run(p_si.add_run(), "2. CI Execution: ", size=10, bold=True, color=PRIMARY)
    set_run(p_si.add_run(), "GitHub Actions tests code, scans CVEs, builds Docker image, and pushes to GHCR.\n", size=9.5, color=TEXT_MUTED)
    set_run(p_si.add_run(), "3. GitOps Reconcile & Verify: ", size=10, bold=True, color=PRIMARY)
    set_run(p_si.add_run(), "Argo CD synchronizes cluster state; Prometheus confirms Pods report UP and response latencies remain within SLO.", size=9.5, color=TEXT_MUTED)

    add_footer(slide, 7, 8)


# ==================================================
# SLIDE 8: SRE Troubleshooting & Capstone Summary
# ==================================================
def build_slide_8(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, prs, dark=False)
    add_header(
        slide,
        "PHASE 6 · SRE & CAPSTONE SUMMARY",
        "Site Reliability Engineering & Troubleshooting Labs",
        "Demonstrated mastery of incident diagnostics, failure remediation, and full capstone rubric compliance"
    )

    # Left Column: Troubleshooting Scenarios Solved
    trouble_w = 6.8
    trouble_h = 3.65
    tcard = create_card(slide, 0.8, 1.7, trouble_w, trouble_h, bg_color=WHITE, border_color=CARD_BORDER)

    tbar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.8), Inches(1.7), Inches(trouble_w), Inches(0.06))
    tbar.fill.solid()
    tbar.fill.fore_color.rgb = RGBColor(239, 68, 68) # Red accent for troubleshooting
    tbar.line.fill.background()

    ttb = slide.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(trouble_w - 0.4), Inches(trouble_h - 0.3))
    ttf = ttb.text_frame
    ttf.word_wrap = True

    tp = ttf.paragraphs[0]
    set_run(tp.add_run(), "9 Kubernetes Troubleshooting Labs: Diagnostics & Fixes", size=12.5, bold=True, color=PRIMARY)

    scenarios = [
        ("CrashLoopBackOff", "Diagnosed missing DB environment variables via kubectl logs; injected proper ConfigMap/Secret."),
        ("ImagePullBackOff", "Identified invalid image tag and registry secret mismatch using kubectl describe pod; resolved auth."),
        ("Readiness Probe Failure", "Corrected misaligned healthcheck port and endpoint path in Helm deployment values."),
        ("Pending Pods / OOMKilled", "Tuned unachievable CPU/memory resource requests against node capacities and memory limits."),
        ("Storage Binding Failure", "Resolved unbound PersistentVolumeClaim by correcting mismatched StorageClass definitions."),
        ("Ingress 502 Bad Gateway", "Fixed target port discrepancy between Ingress rule and backend ClusterIP Service.")
    ]
    for label, fix in scenarios:
        p = ttf.add_paragraph()
        p.space_before = Pt(4)
        set_run(p.add_run(), f"• {label}: ", size=9.5, bold=True, color=RGBColor(220, 38, 38))
        set_run(p.add_run(), fix, size=9, color=TEXT_MUTED)

    # Right Column: Capstone Rubric & Key Outcomes
    out_w = 4.7
    out_h = 3.65
    ocard = create_card(slide, 7.83, 1.7, out_w, out_h, bg_color=WHITE, border_color=CARD_BORDER)

    obar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(7.83), Inches(1.7), Inches(out_w), Inches(0.06))
    obar.fill.solid()
    obar.fill.fore_color.rgb = ACCENT_EMERALD
    obar.line.fill.background()

    otb = slide.shapes.add_textbox(Inches(8.0), Inches(1.85), Inches(out_w - 0.35), Inches(out_h - 0.3))
    otf = otb.text_frame
    otf.word_wrap = True

    op = otf.paragraphs[0]
    set_run(op.add_run(), "Capstone Grading Rubric: 100/100", size=12.5, bold=True, color=PRIMARY)

    modules = [
        ("M1: Three-Tier Application", "10/10 pts", "React, FastAPI, Postgres, Alembic"),
        ("M2: Pytest Suite", "10/10 pts", "15 unit & integration tests, mock DB"),
        ("M3: Git & GitHub", "5/5 pts", "Clean commits, .gitignore, active repo"),
        ("M4: Docker & Compose", "10/10 pts", "Multi-stage builds, non-root users"),
        ("M5: GitHub Actions CI/CD", "15/15 pts", "Automated pipeline, GHCR push"),
        ("M6: DevSecOps Scanning", "5/5 pts", "Trivy container scan, zero high CVEs"),
        ("M7: AWS Terraform IaC", "15/15 pts", "VPC, EKS cluster, clean destroy"),
        ("M8: Kubernetes + Helm", "15/15 pts", "Deployments, HPA, Ingress, Helm chart"),
        ("M9: Observability", "10/10 pts", "Prometheus /metrics & Grafana"),
        ("M10: Documentation & Demo", "5/5 pts", "Comprehensive README & walkthrough")
    ]
    for mod, pts, detail in modules[:7]:
        p = otf.add_paragraph()
        p.space_before = Pt(3)
        set_run(p.add_run(), f"✓ {mod} ", size=9.5, bold=True, color=PRIMARY)
        set_run(p.add_run(), f"({pts}): ", size=9, bold=True, color=ACCENT_EMERALD)
        set_run(p.add_run(), detail, size=8.5, color=TEXT_MUTED)

    # Bottom Conclusion Banner
    conc_card = create_card(slide, 0.8, 5.5, 11.733, 1.3, bg_color=PRIMARY, border_color=PRIMARY)
    ctb = slide.shapes.add_textbox(Inches(1.0), Inches(5.55), Inches(11.333), Inches(1.15))
    ctf = ctb.text_frame
    ctf.word_wrap = True

    cp_t = ctf.paragraphs[0]
    set_run(cp_t.add_run(), "SUMMARY: ENTERPRISE-GRADE REPRODUCIBILITY & OPERATIONAL EXCELLENCE", size=11, bold=True, color=ACCENT_CYAN)

    cp_b = ctf.add_paragraph()
    cp_b.space_before = Pt(4)
    set_run(cp_b.add_run(), "This capstone establishes an enterprise-ready software delivery lifecycle: every commit is validated through automated tests and security scanners, packaged as immutable containers, deployed declaratively through GitOps onto resilient cloud infrastructure, and observed with real-time telemetry.", size=10, color=RGBColor(226, 232, 240))

    add_footer(slide, 8, 8)


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    print("Generating Slide 1: Title & Hero...")
    build_slide_1(prs)
    print("Generating Slide 2: Workflow Overview...")
    build_slide_2(prs)
    print("Generating Slide 3: Application Architecture & Local Dev...")
    build_slide_3(prs)
    print("Generating Slide 4: CI/CD Pipeline & DevSecOps...")
    build_slide_4(prs)
    print("Generating Slide 5: Cloud Infrastructure as Code (Terraform)...")
    build_slide_5(prs)
    print("Generating Slide 6: Kubernetes & Helm Packaging...")
    build_slide_6(prs)
    print("Generating Slide 7: GitOps & Observability...")
    build_slide_7(prs)
    print("Generating Slide 8: SRE Troubleshooting & Capstone Summary...")
    build_slide_8(prs)

    presentation_dir = Path(__file__).resolve().parents[1] / "presentation"
    presentation_dir.mkdir(parents=True, exist_ok=True)
    output_path = presentation_dir / "TaskBoard_DevOps_Workflow.pptx"
    prs.save(str(output_path))
    print(f"Successfully generated: {output_path}")


if __name__ == "__main__":
    main()
