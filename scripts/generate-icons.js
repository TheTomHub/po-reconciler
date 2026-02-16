/**
 * Generate add-in icons at all required sizes.
 * Design: Deep blue rounded square with two overlapping document outlines
 * and a green checkmark — communicating "document reconciliation verified."
 *
 * Run: node scripts/generate-icons.js
 */
const Jimp = require("jimp");
const path = require("path");

const SIZES = [16, 32, 64, 128];

// Colors
const BG_COLOR = 0x1A5276FF;       // Deep professional blue
const DOC_BACK = 0xFFFFFF90;       // White, slightly transparent (back doc)
const DOC_FRONT = 0xFFFFFFCC;      // White, more opaque (front doc)
const DOC_LINE = 0xFFFFFF40;       // Subtle lines inside docs
const CHECK_COLOR = 0x27AE60FF;    // Green checkmark
const CHECK_OUTLINE = 0x1E8449FF;  // Darker green for checkmark depth

async function generateIcon(size) {
  const img = new Jimp(size, size, 0x00000000); // transparent

  const pad = Math.max(1, Math.round(size * 0.06));
  const radius = Math.round(size * 0.18);

  // Draw rounded rectangle background
  drawRoundedRect(img, pad, pad, size - pad * 2, size - pad * 2, radius, BG_COLOR);

  if (size <= 16) {
    // At 16px, just draw a simple checkmark on blue background
    drawCheckmark(img, size, CHECK_COLOR, 0.25, 0.75);
  } else {
    // Draw two document rectangles (back slightly offset)
    const docW = Math.round(size * 0.35);
    const docH = Math.round(size * 0.45);

    // Back document (slightly offset up-left)
    const bx = Math.round(size * 0.18);
    const by = Math.round(size * 0.15);
    drawRect(img, bx, by, docW, docH, DOC_BACK);
    // Lines inside back doc
    if (size >= 64) {
      for (let l = 0; l < 3; l++) {
        const ly = by + Math.round(docH * 0.25) + l * Math.round(docH * 0.2);
        const lw = Math.round(docW * 0.7);
        drawRect(img, bx + Math.round(docW * 0.15), ly, lw, Math.max(1, Math.round(size * 0.015)), DOC_LINE);
      }
    }

    // Front document (offset down-right, on top)
    const fx = Math.round(size * 0.30);
    const fy = Math.round(size * 0.25);
    drawRect(img, fx, fy, docW, docH, DOC_FRONT);
    // Lines inside front doc
    if (size >= 64) {
      for (let l = 0; l < 3; l++) {
        const ly = fy + Math.round(docH * 0.25) + l * Math.round(docH * 0.2);
        const lw = Math.round(docW * 0.7);
        drawRect(img, fx + Math.round(docW * 0.15), ly, lw, Math.max(1, Math.round(size * 0.015)), DOC_LINE);
      }
    }

    // Green checkmark (bottom-right area, overlapping documents)
    drawCheckmark(img, size, CHECK_COLOR, 0.45, 0.85);

    // Add a small circle behind checkmark for emphasis at larger sizes
    if (size >= 64) {
      const cx = Math.round(size * 0.65);
      const cy = Math.round(size * 0.65);
      const cr = Math.round(size * 0.18);
      drawFilledCircle(img, cx, cy, cr, CHECK_COLOR);
      // White checkmark on the green circle
      drawCheckmark(img, size, 0xFFFFFFFF, 0.50, 0.80);
    }
  }

  const outPath = path.join(__dirname, "..", "assets", `icon-${size}.png`);
  await img.writeAsync(outPath);
  console.log(`Created icon-${size}.png`);
}

function drawRoundedRect(img, x, y, w, h, r, color) {
  // Fill main body
  drawRect(img, x + r, y, w - 2 * r, h, color);
  drawRect(img, x, y + r, w, h - 2 * r, color);

  // Fill corners with circles
  drawFilledCircle(img, x + r, y + r, r, color);
  drawFilledCircle(img, x + w - r - 1, y + r, r, color);
  drawFilledCircle(img, x + r, y + h - r - 1, r, color);
  drawFilledCircle(img, x + w - r - 1, y + h - r - 1, r, color);
}

function drawRect(img, x, y, w, h, color) {
  for (let dy = 0; dy < h; dy++) {
    for (let dx = 0; dx < w; dx++) {
      const px = x + dx;
      const py = y + dy;
      if (px >= 0 && px < img.bitmap.width && py >= 0 && py < img.bitmap.height) {
        blendPixel(img, px, py, color);
      }
    }
  }
}

function drawFilledCircle(img, cx, cy, r, color) {
  for (let dy = -r; dy <= r; dy++) {
    for (let dx = -r; dx <= r; dx++) {
      if (dx * dx + dy * dy <= r * r) {
        const px = cx + dx;
        const py = cy + dy;
        if (px >= 0 && px < img.bitmap.width && py >= 0 && py < img.bitmap.height) {
          blendPixel(img, px, py, color);
        }
      }
    }
  }
}

function drawCheckmark(img, size, color, fromFrac, toFrac) {
  // Checkmark: two lines forming a V-shape tilted right
  // Bottom of check (the dip)
  const midX = Math.round(size * ((fromFrac + toFrac) / 2 - 0.05));
  const midY = Math.round(size * 0.78);
  // Left start
  const leftX = Math.round(size * fromFrac);
  const leftY = Math.round(size * 0.58);
  // Right end (tip going up)
  const rightX = Math.round(size * toFrac);
  const rightY = Math.round(size * 0.38);

  const thickness = Math.max(2, Math.round(size * 0.08));

  drawThickLine(img, leftX, leftY, midX, midY, thickness, color);
  drawThickLine(img, midX, midY, rightX, rightY, thickness, color);
}

function drawThickLine(img, x0, y0, x1, y1, thickness, color) {
  const dx = x1 - x0;
  const dy = y1 - y0;
  const steps = Math.max(Math.abs(dx), Math.abs(dy), 1);
  const half = thickness / 2;

  for (let i = 0; i <= steps; i++) {
    const cx = Math.round(x0 + (dx * i) / steps);
    const cy = Math.round(y0 + (dy * i) / steps);

    for (let oy = -Math.ceil(half); oy <= Math.ceil(half); oy++) {
      for (let ox = -Math.ceil(half); ox <= Math.ceil(half); ox++) {
        if (ox * ox + oy * oy <= half * half + 1) {
          const px = cx + ox;
          const py = cy + oy;
          if (px >= 0 && px < img.bitmap.width && py >= 0 && py < img.bitmap.height) {
            blendPixel(img, px, py, color);
          }
        }
      }
    }
  }
}

function blendPixel(img, x, y, color) {
  const srcA = (color & 0xFF) / 255;
  if (srcA >= 0.99) {
    img.setPixelColor(color >>> 0, x, y);
    return;
  }
  if (srcA <= 0.01) return;

  const existing = img.getPixelColor(x, y);
  const dstR = (existing >> 24) & 0xFF;
  const dstG = (existing >> 16) & 0xFF;
  const dstB = (existing >> 8) & 0xFF;
  const dstA = (existing & 0xFF) / 255;

  const srcR = (color >> 24) & 0xFF;
  const srcG = (color >> 16) & 0xFF;
  const srcB = (color >> 8) & 0xFF;

  const outA = srcA + dstA * (1 - srcA);
  if (outA === 0) return;

  const outR = Math.round((srcR * srcA + dstR * dstA * (1 - srcA)) / outA);
  const outG = Math.round((srcG * srcA + dstG * dstA * (1 - srcA)) / outA);
  const outB = Math.round((srcB * srcA + dstB * dstA * (1 - srcA)) / outA);

  const result = (((outR & 0xFF) << 24) | ((outG & 0xFF) << 16) | ((outB & 0xFF) << 8) | (Math.round(outA * 255) & 0xFF)) >>> 0;
  img.setPixelColor(result, x, y);
}

async function main() {
  for (const size of SIZES) {
    await generateIcon(size);
  }
  console.log("\nAll icons generated in assets/");
}

main().catch(console.error);
