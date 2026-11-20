#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Translator Service
Handles image recognition and translation via an OpenAI-compatible API.
"""

import base64
import json
import time
from typing import Dict, List
from pathlib import Path

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

import config

_IMAGE_MIME = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.webp': 'image/webp',
    '.bmp': 'image/bmp',
}


class Translator:
    """Handle image translation using an OpenAI-compatible vision model."""

    def __init__(self):
        self.api_key = config.API_KEY
        self.base_url = config.BASE_URL
        self.model = config.MODEL_ID
        self.client = None

        if OpenAI and self.api_key and self.base_url and self.model:
            self.client = OpenAI(api_key=self.api_key, base_url=self.base_url)

    def _require_client(self):
        if OpenAI is None:
            raise ImportError("openai is not installed. Run: pip install openai")
        missing = [
            name for name, value in (
                ("BASE_URL", self.base_url),
                ("API_KEY", self.api_key),
                ("MODEL_ID", self.model),
            ) if not value
        ]
        if missing:
            raise RuntimeError(
                "Missing "
                + ", ".join(missing)
                + ". Set them in .env (see .env.example)."
            )
        if self.client is None:
            self.client = OpenAI(api_key=self.api_key, base_url=self.base_url)

    def encode_image(self, image_path: str) -> str:
        """Encode image to base64."""
        with open(image_path, 'rb') as f:
            return base64.b64encode(f.read()).decode('utf-8')

    def image_data_url(self, image_path: str) -> str:
        """Build a data URL accepted by OpenAI-compatible vision APIs."""
        mime = _IMAGE_MIME.get(Path(image_path).suffix.lower(), 'image/jpeg')
        return f"data:{mime};base64,{self.encode_image(image_path)}"

    def translate_image(self, image_path: str, source_lang: str = 'zh', target_lang: str = 'en') -> Dict:
        """
        Translate content from an image.

        Args:
            image_path: Path to the image file
            source_lang: Source language code
            target_lang: Target language code

        Returns:
            Dictionary containing extracted and translated content
        """
        self._require_client()

        # Build prompt based on language pair
        prompt = self._build_translation_prompt(source_lang, target_lang)

        # Build content
        content = [
            {
                "type": "image_url",
                "image_url": {"url": self.image_data_url(image_path)}
            },
            {
                "type": "text",
                "text": prompt
            }
        ]

        try:
            # Call API
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": content}]
            )

            result_text = response.choices[0].message.content

            # Parse result
            return self._parse_translation_result(result_text, source_lang, target_lang)

        except Exception as e:
            return {
                'error': str(e),
                'original': '',
                'translated': '',
                'fields': []
            }

    def _build_translation_prompt(self, source_lang: str, target_lang: str) -> str:
        """Build translation prompt based on language pair."""
        lang_names = {
            'zh': 'Chinese',
            'en': 'English',
            'ja': 'Japanese',
            'ko': 'Korean',
            'fr': 'French',
            'de': 'German',
            'es': 'Spanish',
            'ru': 'Russian'
        }

        source_name = lang_names.get(source_lang, source_lang)
        target_name = lang_names.get(target_lang, target_lang)

        prompt = f"""Analyze this document image and provide a translation from {source_name} to {target_name}.

Please identify:
1. The document type (passport, ID card, marriage certificate, work certificate, bank statement, etc.)
2. All visible text fields and their values
3. Translate each field name and value to {target_name}

Return the result in the following JSON format:
{{
    "document_type": "Short document name in {target_name} only, such as ID Card, Passport, Household Register, Marriage Certificate, Birth Certificate, Employment Certificate, or Bank Statement",
    "document_type_original": "Document type in original language",
    "fields": [
        {{
            "label_original": "Field label in original language",
            "value_original": "Field value in original language",
            "label_translated": "Field label in {target_name}",
            "value_translated": "Field value in {target_name}"
        }}
    ],
    "full_text_original": "Complete original text extracted",
    "full_text_translated": "Complete translated text"
}}

Important:
- Extract ALL visible text accurately
- Maintain the original structure and format
- For dates, use standard international format (DD MMM YYYY)
- For names, keep original spelling and add transliteration if needed
- For addresses, provide both original and transliterated/translated version
- If any text is unclear, mark as [unclear]

Only return the JSON, no other text."""

        return prompt

    def _parse_translation_result(self, result_text: str, source_lang: str, target_lang: str) -> Dict:
        """Parse translation result from API response."""
        try:
            # Try to extract JSON from response
            if '{' in result_text and '}' in result_text:
                start = result_text.index('{')
                end = result_text.rindex('}') + 1
                json_str = result_text[start:end]
                result = json.loads(json_str)
                return result
        except:
            pass

        # If JSON parsing fails, return raw text
        return {
            'document_type': 'Unknown',
            'document_type_original': 'Unknown',
            'fields': [],
            'full_text_original': '',
            'full_text_translated': result_text,
            'raw_response': result_text
        }

    def batch_translate(self, image_paths: List[str], source_lang: str = 'zh', target_lang: str = 'en') -> List[Dict]:
        """
        Translate multiple images.

        Args:
            image_paths: List of image paths
            source_lang: Source language code
            target_lang: Target language code

        Returns:
            List of translation results
        """
        results = []
        for i, image_path in enumerate(image_paths):
            print(f"Translating image {i+1}/{len(image_paths)}: {Path(image_path).name}")
            result = self.translate_image(image_path, source_lang, target_lang)
            result['image_path'] = image_path
            result['index'] = i
            results.append(result)

            # Add small delay to avoid rate limiting
            if i < len(image_paths) - 1:
                time.sleep(0.5)

        return results

    def analyze_document_type(self, image_path: str) -> str:
        """
        Analyze and identify document type.

        Args:
            image_path: Path to the image

        Returns:
            Document type string
        """
        self._require_client()

        prompt = """Identify the type of this document. Return only one of these types:
- passport
- id_card
- marriage_certificate
- work_certificate
- bank_statement
- birth_certificate
- diploma
- other

Return only the type name, nothing else."""

        content = [
            {"type": "image_url", "image_url": {"url": self.image_data_url(image_path)}},
            {"type": "text", "text": prompt}
        ]

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": content}]
            )
            return response.choices[0].message.content.strip().lower()
        except:
            return 'unknown'
