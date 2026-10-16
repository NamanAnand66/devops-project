# TaskBoard DevOps Workflow Presentation

This folder contains the official 8-slide executive presentation for the TaskBoard DevOps Capstone Project.

## File Information

* **Presentation File**: `TaskBoard_DevOps_Workflow.pptx`
* **Format**: 16:9 Widescreen (`13.333"` × `7.5"`)
* **Slides**: 8 structured pages explaining the end-to-end delivery workflow
* **Author**: Naman Anand (Roll No: 10417)
* **Institution**: Scaler School of Technology
* **Repository**: [github.com/NamanAnand66/devops-project](https://github.com/NamanAnand66/devops-project)

---

## Slide Outline

1. **Slide 1 — Executive Title & Capstone Hero**: Dark obsidian theme (`#0F172A`), category badge, 3 core architecture capability cards (CI/CD, IaC, GitOps), and student author metadata.
2. **Slide 2 — 5-Phase End-to-End Delivery Pipeline**: Sequential process flow (`01 Dev Loop` → `02 CI & Security` → `03 Artifact Registry` → `04 Cloud Infra` → `05 GitOps & SRE`) with architectural principles.
3. **Slide 3 — Application Architecture & Local Dev Loop**: Three-tier architecture (React 19 + Vite frontend, FastAPI async backend with `/health`, `/ready`, `/metrics`, PostgreSQL 16 persistence), and `docker-compose.yml` health-gated local loop.
4. **Slide 4 — Continuous Integration & DevSecOps Security Gates**: 4-stage pipeline with flake8 linting, 15 Pytests (100% route coverage), Bandit SAST, Gitleaks secrets, lean multi-stage non-root Docker builds, and Trivy CVE vulnerability gate before GHCR push.
5. **Slide 5 — Terraform Architecture on AWS**: Multi-AZ VPC network topology (public/private subnets, NAT Gateways), managed Amazon EKS cluster with IRSA and KMS encryption, state locking, and FinOps clean teardown (`terraform destroy`).
6. **Slide 6 — Kubernetes Workload Orchestration & Helm**: Multi-replica Deployments, StatefulSet with EBS PVC, ConfigMaps & Secrets, 3-tier health probes, HPA autoscaling (50% CPU), and custom `taskboard` Helm chart.
7. **Slide 7 — GitOps Delivery (Argo CD) & Full-Stack Observability**: Pull-based declarative continuous delivery (zero cluster credentials in CI), Prometheus metrics scraping, and Grafana real-time RPS/latency/error rate dashboards.
8. **Slide 8 — SRE Incident Troubleshooting & Capstone Summary**: Detailed diagnostics and resolutions for the 9 deliberate Kubernetes failure scenarios (CrashLoopBackOff, ImagePullBackOff, probe failures, OOMKilled, PVC binding, Ingress 502) and 100/100 grading rubric compliance.

---

## How to Regenerate

To regenerate this presentation from source:

```bash
python3 scripts/generate_devops_presentation.py
```
