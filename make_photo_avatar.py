#!/usr/bin/env python3
"""Turn the user's supplied photo (original_photo.jpg) into a clean circular
avatar (photo.png) for the CV sidebar: crop to head-and-shoulders, apply a
circular mask, and add a thin white ring."""

from PIL import Image, ImageDraw

SRC = "original_photo.jpg"
OUT = "photo.png"

img = Image.open(SRC).convert("RGB")
w, h = img.size  # ~689 x 886
print("source size:", w, h)

# Crop a square centred on the face/upper torso.
# Face is roughly centred horizontally, head near the top third.
crop_size = min(w, int(h * 0.72))
cx = w // 2
cy = int(h * 0.40)  # bias upward toward the face
half = crop_size // 2
left = max(0, cx - half)
top = max(0, cy - half)
right = min(w, left + crop_size)
bottom = min(h, top + crop_size)
# adjust if we hit edges
left = right - crop_size if right - left < crop_size else left
top = bottom - crop_size if bottom - top < crop_size else top
img = img.crop((left, top, right, bottom))

# Upscale/normalize to a working size
SIZE = 800
img = img.resize((SIZE, SIZE), Image.LANCZOS)

# Build circular RGBA result with white ring
result = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(result)

# White ring (outer circle)
draw.ellipse([0, 0, SIZE - 1, SIZE - 1], fill=(255, 255, 255, 255))

# Circular mask for the photo (inner circle, leaving ring)
ring = 22
mask = Image.new("L", (SIZE, SIZE), 0)
mdraw = ImageDraw.Draw(mask)
mdraw.ellipse([ring, ring, SIZE - 1 - ring, SIZE - 1 - ring], fill=255)

result.paste(img, (0, 0), mask)

# Downscale for anti-aliasing
result = result.resize((420, 420), Image.LANCZOS)
result.save(OUT)
print("Saved", OUT)
