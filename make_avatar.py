#!/usr/bin/env python3
"""Generate a clean circular avatar placeholder (initials) for the CV.

A real photograph is not used because no photo was provided and it would be
inappropriate to insert a real person's image. This produces a professional
initials-based avatar that sits nicely on the navy sidebar.
"""

from PIL import Image, ImageDraw, ImageFont

SIZE = 800            # working resolution (downscaled for crispness)
INITIALS = "SD"

# Palette (matches the CV)
CIRCLE_BG = (220, 228, 240, 255)   # soft light blue-grey
INITIAL_COLOR = (22, 51, 91, 255)  # navy
RING_COLOR = (255, 255, 255, 255)  # white ring

LATO_BOLD = (
    "/opt/toolchains/.local/share/mise/installs/ruby/3.4.4/lib/ruby/3.4.0/"
    "rdoc/generator/template/darkfish/fonts/Lato-Regular.ttf"
)

img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# White outer ring
draw.ellipse([0, 0, SIZE - 1, SIZE - 1], fill=RING_COLOR)
# Inner colored circle
pad = 24
draw.ellipse([pad, pad, SIZE - 1 - pad, SIZE - 1 - pad], fill=CIRCLE_BG)

# Initials
try:
    font = ImageFont.truetype(LATO_BOLD, 320)
except OSError:
    font = ImageFont.load_default()

bbox = draw.textbbox((0, 0), INITIALS, font=font)
tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
tx = (SIZE - tw) / 2 - bbox[0]
ty = (SIZE - th) / 2 - bbox[1]
draw.text((tx, ty), INITIALS, font=font, fill=INITIAL_COLOR)

# Downscale for anti-aliasing
img = img.resize((400, 400), Image.LANCZOS)
img.save("avatar.png")
print("Saved avatar.png")
