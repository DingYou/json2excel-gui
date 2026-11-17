const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const script = path.join(__dirname, '../scripts/build-macos.sh');
const posix = { skip: process.platform === 'win32' };

// Contract tests use isolated fake prerequisites; a real signed macOS build is a
// separate verification step. These fixtures never build, install, or sign apps.
function fixture(t, extra = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'json2excel-build-contract-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const bin = path.join(dir, 'bin'); fs.mkdirSync(bin);
  const lib = path.join(dir, 'target-lib'); fs.mkdirSync(lib);
  const cache = path.join(dir, 'cache with spaces');
  const log = path.join(dir, 'calls.log');
  const write = (name, content) => fs.writeFileSync(path.join(bin, name), '#!/bin/sh\nset -eu\n' + content, { mode: 0o755 });
  write('uname', 'printf "%s\\n" "${TEST_OS:-Darwin}"\n');
  write('xcode-select', 'test "${TEST_NO_CLT:-0}" = 0\n');
  write('xcrun', 'test "${TEST_NO_CLT:-0}" = 0\n');
  write('rustc', 'if [ "$1" = -vV ]; then printf "host: %s\\n" "${TEST_TARGET:-aarch64-apple-darwin}"; else printf "%s\\n" "$TEST_LIB"; fi\n');
  write('plutil', `case "$2" in
productName|CFBundleExecutable) printf 'json2excel-gui\\n';;
version) printf '1.0.1\\n';;
CFBundleShortVersionString) printf '%s\\n' "\${TEST_VERSION:-1.0.1}";;
identifier|CFBundleIdentifier) printf 'you.ding.json2excel\\n';;
target_directory) cat >/dev/null; printf '%s\\n' "$TEST_CACHE";;
*) exit 1;;
esac\n`);
  write('lipo', 'printf "%s\\n" "${TEST_ARCH:-arm64}"\n');
  write('codesign', `if [ "$1" = --verify ]; then
  printf 'verify|%s\\n' "$*" >> "$TEST_LOG"
  test "\${TEST_BAD_SIGNATURE:-0}" = 0
else printf 'Signature=adhoc\\n'; fi\n`);
  write('cargo', `printf 'cargo|%s\\n' "$*" >> "$TEST_LOG"
case "$1" in
 --version) test "\${TEST_NO_RUST:-0}" = 0; printf 'cargo 1.99.1\\n';;
 tauri)
  if [ "$2" = --version ]; then
    test "\${TEST_NO_CLI:-0}" = 0
    printf '%s\\n' "\${TEST_CLI_VERSION:-tauri-cli 2.10.0}"
  else
    test "$2" = build
    test "$APPLE_SIGNING_IDENTITY" = '-'
    test -z "\${APPLE_ID+x}" && test -z "\${APPLE_CERTIFICATE+x}" && test -z "\${APPLE_API_KEY+x}"
    printf 'cache|%s\\n' "\${CARGO_TARGET_DIR:-}" >> "$TEST_LOG"
    printf 'identity|%s\\n' "$APPLE_SIGNING_IDENTITY" >> "$TEST_LOG"
    mkdir -p "$TEST_CACHE/\${TEST_TARGET:-aarch64-apple-darwin}/release/bundle/macos/json2excel-gui.app/Contents/MacOS"
    printf '{}' > "$TEST_CACHE/\${TEST_TARGET:-aarch64-apple-darwin}/release/bundle/macos/json2excel-gui.app/Contents/Info.plist"
  fi;;
 metadata) printf '{"target_directory":"fixture"}\\n';;
 *) exit 1;;
esac\n`);
  const env = { ...process.env, PATH: `${bin}:/usr/bin:/bin:/usr/sbin:/sbin`, TEST_LIB: lib, TEST_CACHE: cache, TEST_LOG: log, APPLE_ID: 'fixture-only', APPLE_API_KEY: 'fixture-only', APPLE_CERTIFICATE: 'fixture-only', APPLE_SIGNING_IDENTITY: 'fixture-identity', ...extra };
  delete env.CARGO_TARGET_DIR;
  return { dir, cache, log, env, run: (args = [], overrides = {}) => spawnSync('/bin/bash', [script, ...args], { cwd: dir, env: { ...env, ...overrides }, encoding: 'utf8' }), calls: () => fs.existsSync(log) ? fs.readFileSync(log, 'utf8') : '' };
}

test('macOS build script parses and help has no side effects', posix, t => {
  assert.equal(spawnSync('/bin/bash', ['-n', script]).status, 0);
  const f = fixture(t);
  const result = f.run(['--help']);
  assert.equal(result.status, 0); assert.match(result.stdout, /不自动安装/); assert.equal(f.calls(), '');
  assert.doesNotMatch(fs.readFileSync(script, 'utf8'), /\b(?:sudo|xattr|spctl|hdiutil)\b/);
});

test('non-macOS, unavailable Rust and missing CLT stop before any build', posix, t => {
  const f = fixture(t);
  const linux = f.run([], { TEST_OS: 'Linux' });
  assert.notEqual(linux.status, 0); assert.match(linux.stderr, /仅支持 macOS/); assert.equal(f.calls(), '');
  const noRust = f.run([], { TEST_NO_RUST: '1' });
  assert.notEqual(noRust.status, 0); assert.match(noRust.stderr, /Rust\/Cargo 工具链不可用/);
  const noClt = f.run([], { TEST_NO_CLT: '1' });
  assert.notEqual(noClt.status, 0); assert.match(noClt.stderr, /xcode-select --install/); assert.doesNotMatch(f.calls(), /tauri build/);
});

test('missing or incompatible Cargo CLI gives manual instructions, never installs it', posix, t => {
  const f = fixture(t);
  const absent = f.run([], { TEST_NO_CLI: '1' });
  assert.notEqual(absent.status, 0); assert.match(absent.stderr, /cargo install tauri-cli --version 2\.10\.0 --locked/);
  const old = f.run([], { TEST_CLI_VERSION: 'tauri-cli 1.6.0' });
  assert.notEqual(old.status, 0); assert.match(old.stderr, /需要 Cargo Tauri CLI v2/);
  assert.doesNotMatch(f.calls(), /cargo\|install|cargo\|tauri build/);
});

test('locked local app build preserves caller cache path with spaces and scopes signing credentials', posix, t => {
  const f = fixture(t);
  const result = f.run([], { CARGO_TARGET_DIR: 'cache with spaces' });
  assert.equal(result.status, 0, result.stderr);
  assert.match(f.calls(), /cargo\|metadata .*--locked/);
  assert.match(f.calls(), /cargo\|tauri build --ci --bundles app --target aarch64-apple-darwin -- --locked/);
  const recordedCache = f.calls().split('\n').find(line => line.startsWith('cache|')).slice(6);
  assert.equal(fs.realpathSync(recordedCache), fs.realpathSync(f.cache));
  assert.match(f.calls(), /identity\|-/);
  assert.match(f.calls(), /verify\|--verify --deep --strict/);
  assert.equal(result.stdout.trim(), path.join(f.cache, 'aarch64-apple-darwin/release/bundle/macos/json2excel-gui.app'));
  assert.equal(f.env.APPLE_SIGNING_IDENTITY, 'fixture-identity');
  assert.doesNotMatch(f.calls(), /--bundles dmg|cargo\|install|\/Applications/);
});

test('Intel target is supported explicitly; unsupported or missing Rust targets fail', posix, t => {
  const f = fixture(t);
  const intel = f.run(['x86_64-apple-darwin'], { TEST_TARGET: 'x86_64-apple-darwin', TEST_ARCH: 'x86_64' });
  assert.equal(intel.status, 0, intel.stderr);
  const unknown = f.run(['universal-apple-darwin']);
  assert.notEqual(unknown.status, 0); assert.match(unknown.stderr, /不支持目标/);
  const missing = f.run([], { TEST_LIB: path.join(f.dir, 'absent') });
  assert.notEqual(missing.status, 0); assert.match(missing.stderr, /rustup target add/);
});

test('architecture, version and signature mismatches reject a stale/unusable bundle', posix, t => {
  const f = fixture(t);
  for (const [env, message] of [[{ TEST_ARCH: 'x86_64' }, /产物架构/], [{ TEST_VERSION: '0.0.0' }, /产物版本/], [{ TEST_BAD_SIGNATURE: '1' }, /本机自用构建/]]) {
    const result = f.run([], env);
    assert.notEqual(result.status, 0); assert.match(result.stderr, message);
    assert.equal(result.stdout, '');
  }
});
