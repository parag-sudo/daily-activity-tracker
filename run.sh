#!/bin/bash
cd "$(dirname "$0")"
pip install -r requirements.txt -q
SMTP_PASS="your_gmail_app_password_here" python3 app.py
