/**
 * Generates RichCell PNG icons (16×16, 32×32, 80×80).
 * Design: spreadsheet cell grid + paintbrush in top-right corner.
 *
 * Usage:  node generate-icons.js
 * Needs:  sharp  (npm install)
 */

'use strict';

const sharp = require('sharp');
const fs    = require('fs');
const path  = require('path');

const ASSETS_DIR = path.join(__dirname, 'assets');
if (!fs.existsSync(ASSETS_DIR)) fs.mkdirSync(ASSETS_DIR);

/**
 * Returns an SVG string for the RichCell icon at the given size.
 * All coordinates are in a 80×80 viewBox and scale automatically.
 */
function buildSVG() {
  return `<svg viewBox="0 0 80 80" xmlns="http://www.w3.org/2000/svg">

    <!-- ── Cell grid ─────────────────────────────────── -->
    <!-- Outer border -->
    <rect x="2" y="18" width="54" height="60" rx="5"
          fill="white" stroke="#217346" stroke-width="3"/>

    <!-- Horizontal dividers -->
    <line x1="2"  y1="38" x2="56" y2="38" stroke="#217346" stroke-width="2"/>
    <line x1="2"  y1="58" x2="56" y2="58" stroke="#217346" stroke-width="2"/>

    <!-- Vertical divider -->
    <line x1="28" y1="18" x2="28" y2="78" stroke="#217346" stroke-width="2"/>

    <!-- Content lines (suggest text in cells) -->
    <line x1="7"  y1="49" x2="22" y2="49" stroke="#c0d9c8" stroke-width="2.5" stroke-linecap="round"/>
    <line x1="33" y1="49" x2="50" y2="49" stroke="#c0d9c8" stroke-width="2.5" stroke-linecap="round"/>
    <line x1="7"  y1="69" x2="22" y2="69" stroke="#c0d9c8" stroke-width="2.5" stroke-linecap="round"/>
    <line x1="33" y1="69" x2="50" y2="69" stroke="#c0d9c8" stroke-width="2.5" stroke-linecap="round"/>

    <!-- ── Paintbrush ──────────────────────────────────── -->
    <!-- Handle (wood) -->
    <line x1="78" y1="2" x2="50" y2="30"
          stroke="#8B5E3C" stroke-width="6"
          stroke-linecap="round"/>

    <!-- Ferrule (silver band) -->
    <line x1="46" y1="26" x2="54" y2="34"
          stroke="#AAAAAA" stroke-width="8"
          stroke-linecap="butt"/>
    <line x1="47" y1="25" x2="55" y2="33"
          stroke="#CCCCCC" stroke-width="3"
          stroke-linecap="butt" opacity="0.6"/>

    <!-- Bristles (green — painting with Excel green) -->
    <ellipse cx="40" cy="40" rx="8" ry="5"
             fill="#217346"
             transform="rotate(-45 40 40)"/>
    <ellipse cx="36" cy="44" rx="5" ry="3"
             fill="#1a5c38"
             transform="rotate(-45 36 44)"/>

    <!-- Paint drop -->
    <circle cx="31" cy="49" r="3" fill="#217346" opacity="0.55"/>

  </svg>`;
}

async function createIcon(size, filename) {
  const destPath = path.join(ASSETS_DIR, filename);
  await sharp(Buffer.from(buildSVG()))
    .resize(size, size)
    .png()
    .toFile(destPath);
  console.log(`  ✔  ${filename}  (${size}×${size} px)`);
}

(async () => {
  console.log('\nGenerating RichCell icons…');
  await createIcon(16, 'icon-16.png');
  await createIcon(32, 'icon-32.png');
  await createIcon(80, 'icon-80.png');
  console.log('Done.\n');
})();
