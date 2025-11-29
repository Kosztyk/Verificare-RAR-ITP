import asyncio
import io

import pytesseract
from PIL import Image, ImageFilter, ImageOps, ImageStat

DIGITS_ONLY_CFG = "--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789"


def _clean(img: Image.Image) -> Image.Image:
    """Aggressively denoise and threshold the CAPTCHA image for OCR."""

    img = ImageOps.grayscale(img)
    img = ImageOps.autocontrast(img, cutoff=5)
    img = img.filter(ImageFilter.MedianFilter(size=3))
    img = img.filter(ImageFilter.GaussianBlur(radius=0.5))

    # Adaptive threshold around the image mean to handle varying backgrounds
    mean = ImageStat.Stat(img).mean[0]
    threshold = max(100, min(170, int(mean)))
    img = img.point(lambda p: 255 if p > threshold else 0, "1")

    return img.resize((img.width * 2, img.height * 2), Image.LANCZOS)

async def solve_captcha_image(raw: bytes) -> str:
    def _ocr() -> str:
        img = Image.open(io.BytesIO(raw))
        img = _clean(img)
        return pytesseract.image_to_string(img, config=DIGITS_ONLY_CFG).strip()
    return await asyncio.to_thread(_ocr)
