import io, asyncio
from PIL import Image, ImageOps, ImageFilter
import pytesseract

DIGITS_ONLY_CFG = "--psm 8 -c tessedit_char_whitelist=0123456789"

def _clean(img: Image.Image) -> Image.Image:
    img = ImageOps.grayscale(img)
    img = ImageOps.autocontrast(img, cutoff=5)
    img = img.filter(ImageFilter.MedianFilter(3))
    img = img.point(lambda p: 0 if p < 128 else 255, "1")
    return img.resize((img.width*2, img.height*2), Image.LANCZOS)

async def solve_captcha_image(raw: bytes) -> str:
    def _ocr() -> str:
        img = Image.open(io.BytesIO(raw))
        img = _clean(img)
        return pytesseract.image_to_string(img, config=DIGITS_ONLY_CFG).strip()
    return await asyncio.to_thread(_ocr)
