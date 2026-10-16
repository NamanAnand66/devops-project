#!/usr/bin/env python3
"""
Generate TaskBoard DevOps workflow presentation (8 slides, 16:9 widescreen).
Builds a high-impact, professional deck detailing the entire end-to-end delivery lifecycle.
"""

import sys
from pathlib import Path

# Delegate to the comprehensive generator
script_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(script_dir))

from generate_devops_presentation import main

if __name__ == "__main__":
    main()
