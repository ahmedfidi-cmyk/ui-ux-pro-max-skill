#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Portrait / Headshot Retouch via Gemini Nano Banana (image-to-image)

Models:
- Nano Banana (default): gemini-2.5-flash-image
- Nano Banana Pro (--pro): gemini-3-pro-image-preview

Usage:
    python generate.py --input photo.jpg --preset linkedin
    python generate.py --input photo.jpg --preset linkedin --pro --output headshot.png
    python generate.py --input photo.jpg --prompt "custom retouch instructions"
"""

import argparse
import mimetypes
import os
import sys
from datetime import datetime
from pathlib import Path


def load_env():
    env_paths = [
        Path(__file__).parent.parent.parent / ".env",
        Path.home() / ".claude" / "skills" / ".env",
        Path.home() / ".claude" / ".env",
    ]
    for env_path in env_paths:
        if env_path.exists():
            with open(env_path) as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        key, value = line.split("=", 1)
                        if key not in os.environ:
                            os.environ[key] = value.strip("\"'")


load_env()

try:
    from google import genai
    from google.genai import types
except ImportError:
    print("Error: google-genai package not installed.")
    print("Install with: pip install google-genai")
    sys.exit(1)


GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
GEMINI_FLASH = "gemini-2.5-flash-image"
GEMINI_PRO = "gemini-3-pro-image-preview"

ASPECT_RATIOS = ["1:1", "4:5", "3:4", "9:16"]
DEFAULT_ASPECT_RATIO = "1:1"

PRESETS = {
    "linkedin": (
        "Retouch this photograph into a professional LinkedIn headshot. "
        "Preserve the subject's exact identity, face geometry, skin tone, hairline, "
        "facial hair, eye shape and color, and natural expression — do not alter the person. "
        "Keep the traditional Saudi attire intact and accurate: crisp white thobe, "
        "white ghutra (head cloth), and black agal (rope band on top). "
        "Keep the friendly, confident smile and the wristwatch on the left wrist. "
        "Enhance lighting to soft, even, studio-quality with a gentle key light from the front, "
        "subtle fill, and clean catchlights in the eyes. "
        "Apply subtle, realistic skin retouching: even skin tone, reduce minor blemishes, "
        "but keep natural pores, texture, and stubble — no plastic smoothing, no slimming, "
        "no reshaping of facial features. "
        "Replace the background with a smooth, soft corporate gradient — light neutral gray "
        "transitioning to a hint of soft blue — clean and uncluttered, no objects, no text. "
        "Improve overall clarity, contrast, white balance, and color accuracy. "
        "Whites of the thobe and ghutra should read clean white without blowing out. "
        "Center the subject, head-and-shoulders crop suitable for a LinkedIn profile photo, "
        "1:1 square framing with appropriate headroom. "
        "Sharp, high-resolution, photorealistic, modern executive aesthetic. "
        "Do not add watermarks, logos, jewelry, glasses, or any element not present in the source."
    ),
    "corporate": (
        "Retouch into a polished corporate headshot. Preserve identity exactly. "
        "Soft studio lighting, neutral gray gradient background, natural skin texture, "
        "subtle retouching only, 1:1 crop, head-and-shoulders, business-ready, photorealistic."
    ),
    "passport": (
        "Retouch into a passport-style portrait. Preserve identity exactly. "
        "Even flat lighting, plain off-white background, neutral expression preserved if present, "
        "natural skin texture, no shadows on background, sharp focus on face, 1:1 crop."
    ),
}

NEGATIVE_GUARDRAILS = (
    "Do not change the person's identity, race, age, gender, body shape, or facial structure. "
    "Do not generate a different person. Do not add or remove clothing items. "
    "Do not add accessories, jewelry, glasses, makeup, beauty filters, or AI-style smoothing. "
    "Do not add text, captions, watermarks, or borders."
)


def read_image_bytes(path):
    p = Path(path)
    if not p.exists():
        print(f"Error: input image not found: {path}")
        sys.exit(1)
    mime, _ = mimetypes.guess_type(str(p))
    if not mime or not mime.startswith("image/"):
        mime = "image/jpeg"
    return p.read_bytes(), mime


def retouch(input_path, prompt, output_path=None, use_pro=False, aspect_ratio=None):
    if not GEMINI_API_KEY:
        print("Error: GEMINI_API_KEY not set")
        print("Set it with: export GEMINI_API_KEY='your-key'")
        return None

    client = genai.Client(api_key=GEMINI_API_KEY)
    model = GEMINI_PRO if use_pro else GEMINI_FLASH
    model_label = "Nano Banana Pro" if use_pro else "Nano Banana"
    ratio = aspect_ratio if aspect_ratio in ASPECT_RATIOS else DEFAULT_ASPECT_RATIO

    img_bytes, mime = read_image_bytes(input_path)
    full_prompt = f"{prompt}\n\nConstraints: {NEGATIVE_GUARDRAILS}"

    print(f"Retouching with {model_label} ({model})")
    print(f"Input:  {input_path} ({mime}, {len(img_bytes)} bytes)")
    print(f"Aspect: {ratio}")
    print()

    image_part = types.Part.from_bytes(data=img_bytes, mime_type=mime)

    response = client.models.generate_content(
        model=model,
        contents=[image_part, full_prompt],
        config=types.GenerateContentConfig(
            response_modalities=["IMAGE", "TEXT"],
            image_config=types.ImageConfig(aspect_ratio=ratio),
            safety_settings=[
                types.SafetySetting(
                    category="HARM_CATEGORY_HATE_SPEECH",
                    threshold="BLOCK_LOW_AND_ABOVE",
                ),
                types.SafetySetting(
                    category="HARM_CATEGORY_DANGEROUS_CONTENT",
                    threshold="BLOCK_LOW_AND_ABOVE",
                ),
                types.SafetySetting(
                    category="HARM_CATEGORY_SEXUALLY_EXPLICIT",
                    threshold="BLOCK_LOW_AND_ABOVE",
                ),
                types.SafetySetting(
                    category="HARM_CATEGORY_HARASSMENT",
                    threshold="BLOCK_LOW_AND_ABOVE",
                ),
            ],
        ),
    )

    image_data = None
    text_notes = []
    for part in response.candidates[0].content.parts:
        if hasattr(part, "inline_data") and part.inline_data and part.inline_data.mime_type.startswith("image/"):
            image_data = part.inline_data.data
        elif hasattr(part, "text") and part.text:
            text_notes.append(part.text)

    if not image_data:
        print("No image returned by the model.")
        if text_notes:
            print("Model text response:")
            print("\n".join(text_notes))
        return None

    if output_path is None:
        stem = Path(input_path).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"{stem}_retouched_{timestamp}.png"

    Path(output_path).write_bytes(image_data)
    print(f"Saved: {output_path}")
    if text_notes:
        print("Model notes:", " ".join(text_notes)[:300])
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Image-to-image portrait retouch via Gemini Nano Banana")
    parser.add_argument("--input", "-i", required=True, help="Path to input photo")
    parser.add_argument("--preset", choices=list(PRESETS.keys()), help="Built-in retouch preset")
    parser.add_argument("--prompt", "-p", help="Custom retouch instructions (overrides --preset)")
    parser.add_argument("--output", "-o", help="Output path (default: <input>_retouched_<timestamp>.png)")
    parser.add_argument("--pro", action="store_true", help="Use Nano Banana Pro (higher quality)")
    parser.add_argument(
        "--aspect-ratio",
        "-r",
        choices=ASPECT_RATIOS,
        default=DEFAULT_ASPECT_RATIO,
        help=f"Output aspect ratio (default: {DEFAULT_ASPECT_RATIO})",
    )

    args = parser.parse_args()

    if not args.prompt and not args.preset:
        args.preset = "linkedin"

    prompt = args.prompt or PRESETS[args.preset]

    retouch(
        input_path=args.input,
        prompt=prompt,
        output_path=args.output,
        use_pro=args.pro,
        aspect_ratio=args.aspect_ratio,
    )


if __name__ == "__main__":
    main()
