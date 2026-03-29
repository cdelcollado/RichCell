/**
 * Generates simple PNG icon files required by manifest.xml.
 * Each icon is a solid green circle on a transparent background.
 *
 * Usage:  node generate-icons.js
 * Needs:  pngjs  (installed via npm install)
 */

'use strict';

const { PNG } = require('pngjs');
const fs      = require('fs');
const path    = require('path');

const ASSETS_DIR = path.join(__dirname, 'assets');

// Ensure assets/ folder exists
if (!fs.existsSync(ASSETS_DIR)) {
  fs.mkdirSync(ASSETS_DIR);
}

// Icon colour: Excel green  #217346  →  R=33, G=115, B=70
const R = 33, G = 115, B = 70;

/**
 * Creates a SIZE×SIZE PNG with a filled circle centred in the image.
 * Pixels outside the circle are fully transparent.
 */
function createIcon(size, filename) {
  const png = new PNG({ width: size, height: size, filterType: -1 });
  const cx  = (size - 1) / 2;
  const cy  = (size - 1) / 2;
  const r   = size * 0.44;   // radius slightly inset from edge

  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const idx  = (size * y + x) << 2;   // * 4
      const dist = Math.sqrt((x - cx) ** 2 + (y - cy) ** 2);

      if (dist <= r) {
        // Inside circle → solid green
        png.data[idx]     = R;
        png.data[idx + 1] = G;
        png.data[idx + 2] = B;
        png.data[idx + 3] = 255;
      } else {
        // Outside → transparent
        png.data[idx]     = 0;
        png.data[idx + 1] = 0;
        png.data[idx + 2] = 0;
        png.data[idx + 3] = 0;
      }
    }
  }

  const buffer   = PNG.sync.write(png);
  const destPath = path.join(ASSETS_DIR, filename);
  fs.writeFileSync(destPath, buffer);
  console.log(`  ✔  ${filename}  (${size}×${size} px)`);
}

console.log('\nGenerating add-in icons…');
createIcon(16, 'icon-16.png');
createIcon(32, 'icon-32.png');
createIcon(80, 'icon-80.png');
console.log('Done.\n');
