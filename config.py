#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configuration settings for Visa Translator Web Application
"""

import os
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

# Base directory
BASE_DIR = Path(__file__).parent

if load_dotenv is not None:
    load_dotenv(BASE_DIR / ".env")

# Upload settings
UPLOAD_FOLDER = BASE_DIR / "static" / "uploads"
OUTPUT_FOLDER = BASE_DIR / "static" / "outputs"
ALLOWED_EXTENSIONS = {'pdf', 'xlsx', 'xls', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max file size

# OpenAI-compatible chat API (vision models). Set these in .env.
BASE_URL = os.environ.get("BASE_URL", "").strip()
API_KEY = os.environ.get("API_KEY", "").strip()
MODEL_ID = os.environ.get("MODEL_ID", "").strip()

# Supported languages
SUPPORTED_LANGUAGES = {
    'zh': '中文',
    'en': 'English',
    'ja': '日本語',
    'ko': '한국어',
    'fr': 'Français',
    'de': 'Deutsch',
    'es': 'Español',
    'ru': 'Русский'
}

# Language pairs for translation (using string keys for JSON serialization)
LANGUAGE_PAIRS = {
    'zh-en': '中文 → English',
    'zh-ja': '中文 → 日本語',
    'zh-ko': '中文 → 한국어',
    'en-zh': 'English → 中文',
    'en-ja': 'English → 日本語',
    'en-ko': 'English → 한국어',
}

# Flask secret key
SECRET_KEY = os.environ.get("SECRET_KEY", "visa-translator-secret-key-2024")

# Ensure directories exist
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
