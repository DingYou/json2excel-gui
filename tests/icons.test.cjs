const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const zlib = require('node:zlib');
const root = path.join(__dirname, '..');
const pngMagic = Buffer.from('89504e470d0a1a0a', 'hex');

// Read-only validation of the generator's RGBA PNGs, including real alpha samples.
function png(buffer) {
  assert.ok(buffer.subarray(0, 8).equals(pngMagic));
  const chunks = []; let width, height;
  for (let offset = 8; offset < buffer.length;) {
    const length = buffer.readUInt32BE(offset);
    assert.ok(offset + 12 + length <= buffer.length);
    const type = buffer.toString('ascii', offset + 4, offset + 8);
    const data = buffer.subarray(offset + 8, offset + 8 + length);
    if (type === 'IHDR') {
      width = data.readUInt32BE(0); height = data.readUInt32BE(4);
      assert.equal(data[8], 8); assert.equal(data[9], 6); assert.equal(data[12], 0);
    }
    if (type === 'IDAT') chunks.push(data);
    offset += length + 12;
    if (type === 'IEND') { assert.equal(offset, buffer.length); break; }
  }
  const stride = width * 4;
  const raw = zlib.inflateSync(Buffer.concat(chunks), { maxOutputLength: (stride + 1) * height });
  assert.equal(raw.length, (stride + 1) * height);
  let previous = Buffer.alloc(stride), cursor = 0, minAlpha = 255, maxAlpha = 0, farMarginAlpha = 0;
  const paeth = (a, b, c) => {
    const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
    return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
  };
  for (let y = 0; y < height; y++) {
    const filter = raw[cursor++]; assert.ok(filter <= 4);
    const row = Buffer.alloc(stride);
    for (let x = 0; x < stride; x++) {
      const left = x >= 4 ? row[x - 4] : 0, up = previous[x], upperLeft = x >= 4 ? previous[x - 4] : 0;
      const prediction = [0, left, up, Math.floor((left + up) / 2), paeth(left, up, upperLeft)][filter];
      row[x] = (raw[cursor++] + prediction) & 255;
      if (x % 4 === 3) {
        minAlpha = Math.min(minAlpha, row[x]); maxAlpha = Math.max(maxAlpha, row[x]);
        const pixel = Math.floor(x / 4);
        if (pixel < width * 0.06 || pixel >= width * 0.94 || y < height * 0.06 || y >= height * 0.94) farMarginAlpha = Math.max(farMarginAlpha, row[x]);
      }
    }
    previous = row;
  }
  return { width, height, minAlpha, maxAlpha, farMarginAlpha };
}
const read = relative => fs.readFileSync(path.join(root, relative));
function assertSizeAndAlpha(buffer, size) {
  const info = png(buffer);
  assert.equal(info.width, size); assert.equal(info.height, size);
  assert.equal(info.minAlpha, 0); assert.equal(info.maxAlpha, 255);
}

test('canonical icon is the exact owner-selected milky periwinkle PNG, with transparent alpha', () => {
  const source = read('app-icon.png');
  assert.equal(crypto.createHash('sha256').update(source).digest('hex'), '4de24f79619162a1c7e29c47908511a46de1b635efcb11de6369ae0e95032310');
  assertSizeAndAlpha(source, 1254);
  assert.ok(png(source).farMarginAlpha <= 2, 'Far margins remain effectively transparent without altering the selected foreground');
  assert.ok(fs.existsSync(path.join(root, 'docs/legacy/app-icon-v1.svg')));
});

test('all top-level desktop PNG sizes and configured icon paths exist and retain transparency', () => {
  const icons = path.join(root, 'src-tauri/icons');
  for (const filename of fs.readdirSync(icons).filter(name => name.endsWith('.png'))) {
    const size = filename === 'icon.png' ? 512 : filename === 'StoreLogo.png' ? 50 : filename === '128x128@2x.png' ? 256 : Number(filename.match(/^(?:Square)?(\d+)/)?.[1]);
    assert.ok(size > 0, filename); assertSizeAndAlpha(fs.readFileSync(path.join(icons, filename)), size);
  }
  const config = JSON.parse(read('src-tauri/tauri.conf.json'));
  for (const icon of config.bundle.icon) assert.ok(fs.statSync(path.join(root, 'src-tauri', icon)).size > 0, icon);
});

test('Windows ICO contains bounded 32-bit PNG frames including small and 256px sizes', () => {
  const data = read('src-tauri/icons/icon.ico');
  assert.equal(data.readUInt16LE(0), 0); assert.equal(data.readUInt16LE(2), 1);
  const count = data.readUInt16LE(4), sizes = new Set(); assert.ok(count >= 6);
  for (let i = 0; i < count; i++) {
    const entry = 6 + i * 16, size = data[entry] || 256;
    assert.equal(data[entry + 1] || 256, size); assert.equal(data.readUInt16LE(entry + 6), 32);
    const length = data.readUInt32LE(entry + 8), offset = data.readUInt32LE(entry + 12);
    assert.ok(offset >= 6 + count * 16 && offset + length <= data.length);
    assertSizeAndAlpha(data.subarray(offset, offset + length), size); sizes.add(size);
  }
  for (const size of [16, 24, 32, 48, 64, 256]) assert.ok(sizes.has(size));
});

test('macOS ICNS contains intact image chunks through 1024px and small-size alpha masks', () => {
  const data = read('src-tauri/icons/icon.icns');
  assert.equal(data.toString('ascii', 0, 4), 'icns'); assert.equal(data.readUInt32BE(4), data.length);
  const pngSizes = new Set(), masks = new Map();
  for (let offset = 8; offset < data.length;) {
    const length = data.readUInt32BE(offset + 4); assert.ok(length >= 8 && offset + length <= data.length);
    const type = data.toString('ascii', offset, offset + 4), frame = data.subarray(offset + 8, offset + length);
    if (frame.subarray(0, 8).equals(pngMagic)) {
      const info = png(frame); assert.equal(info.width, info.height); assertSizeAndAlpha(frame, info.width); pngSizes.add(info.width);
    }
    if (type === 's8mk' || type === 'l8mk') masks.set(type, frame);
    offset += length;
  }
  for (const size of [32, 64, 128, 256, 512, 1024]) assert.ok(pngSizes.has(size));
  for (const [type, size] of [['s8mk', 16], ['l8mk', 32]]) {
    const mask = masks.get(type); assert.equal(mask.length, size * size); assert.equal(Math.min(...mask), 0); assert.equal(Math.max(...mask), 255);
  }
});
