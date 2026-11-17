#!/bin/bash
# Local source build only. No installer, Node/npm, DMG, or OS-security changes.
set -euo pipefail

fail() { printf '错误：%s\n' "$*" >&2; exit 1; }

if [[ $# -gt 1 || ${1:-} == --help || ${1:-} == -h ]]; then
  printf '用法：bash scripts/build-macos.sh [aarch64-apple-darwin|x86_64-apple-darwin]\n'
  printf '默认使用当前 Rust 工具链的 macOS 架构，仅生成并校验本机自用 .app，不自动安装。\n'
  [[ $# -le 1 ]] && exit 0 || exit 1
fi

[[ $(uname -s) == Darwin ]] || fail '此脚本仅支持 macOS。'
for tool in cargo rustc; do
  command -v "$tool" >/dev/null 2>&1 || fail '缺少 Rust/Cargo；请先按 https://www.rust-lang.org/tools/install 安装。'
done
cargo --version >/dev/null 2>&1 && rustc --version >/dev/null 2>&1 ||
  fail 'Rust/Cargo 工具链不可用；请先手动配置可用的 Rust 工具链。'
for tool in xcode-select xcrun; do
  command -v "$tool" >/dev/null 2>&1 || fail '缺少 Xcode Command Line Tools；请手动运行 xcode-select --install。'
done
xcode-select -p >/dev/null 2>&1 && xcrun --find clang >/dev/null 2>&1 ||
  fail 'Xcode Command Line Tools 未就绪；请先手动完成 xcode-select --install。'
for tool in plutil codesign lipo; do
  command -v "$tool" >/dev/null 2>&1 || fail "缺少 macOS 系统工具 ${tool}；请检查本机开发环境。"
done
if ! cli_version=$(cargo tauri --version 2>/dev/null); then
  fail '缺少 Cargo Tauri CLI；请手动运行 cargo install tauri-cli --version 2.10.0 --locked。脚本不会自动安装工具。'
fi
case "$cli_version" in
  'tauri-cli 2.'*) ;;
  *) fail "需要 Cargo Tauri CLI v2（实测 2.10.0），当前为：$cli_version" ;;
esac

ROOT=$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd -P)
# Preserve the caller's relative cache location before changing into the repository.
if [[ -n ${CARGO_TARGET_DIR:-} && $CARGO_TARGET_DIR != /* ]]; then
  export CARGO_TARGET_DIR="$PWD/$CARGO_TARGET_DIR"
fi
cd "$ROOT"
TARGET=${1:-$(rustc -vV | awk '$1 == "host:" { print $2 }')}
case "$TARGET" in
  aarch64-apple-darwin) ARCH=arm64 ;;
  x86_64-apple-darwin) ARCH=x86_64 ;;
  *) fail "不支持目标 ${TARGET}；只支持 aarch64-apple-darwin 或 x86_64-apple-darwin。" ;;
esac
TARGET_LIB=$(rustc --print target-libdir --target "$TARGET")
[[ -d "$TARGET_LIB" ]] || fail "Rust 目标 ${TARGET} 未安装；若使用 rustup，请手动运行 rustup target add ${TARGET}。"

CONFIG="$ROOT/src-tauri/tauri.conf.json"
PRODUCT=$(plutil -extract productName raw -o - "$CONFIG")
VERSION=$(plutil -extract version raw -o - "$CONFIG")
IDENTIFIER=$(plutil -extract identifier raw -o - "$CONFIG")
METADATA=$(cargo metadata --manifest-path "$ROOT/src-tauri/Cargo.toml" --no-deps --format-version 1 --locked)
TARGET_DIR=$(printf '%s' "$METADATA" | plutil -extract target_directory raw -o - -)
[[ "$TARGET_DIR" = /* ]] || fail 'Cargo metadata 未返回绝对 target_directory。'
APP="$TARGET_DIR/$TARGET/release/bundle/macos/$PRODUCT.app"

printf '本机自用构建：%s；%s；缓存目录：%s\n' "$cli_version" "$TARGET" "$TARGET_DIR" >&2
# Do not import certificates or submit notarization from a local self-use build.
# All environment overrides apply only to this subprocess, never the user's shell.
env -u APPLE_CERTIFICATE -u APPLE_CERTIFICATE_PASSWORD \
  -u APPLE_ID -u APPLE_PASSWORD -u APPLE_TEAM_ID \
  -u APPLE_API_KEY -u APPLE_API_ISSUER -u APPLE_API_KEY_PATH \
  -u APPLE_KEYCHAIN -u APPLE_KEYCHAIN_PROFILE \
  APPLE_SIGNING_IDENTITY=- \
  cargo tauri build --ci --bundles app --target "$TARGET" -- --locked

[[ -d "$APP" ]] || fail "未找到预期 .app：${APP}（请检查额外平台配置是否更改了 productName）。"
EXECUTABLE=$(plutil -extract CFBundleExecutable raw -o - "$APP/Contents/Info.plist")
[[ $(lipo -archs "$APP/Contents/MacOS/$EXECUTABLE") == "$ARCH" ]] || fail '产物架构与请求目标不一致。'
[[ $(plutil -extract CFBundleShortVersionString raw -o - "$APP/Contents/Info.plist") == "$VERSION" ]] || fail '产物版本与配置不一致。'
[[ $(plutil -extract CFBundleIdentifier raw -o - "$APP/Contents/Info.plist") == "$IDENTIFIER" ]] || fail '产物标识与配置不一致。'
codesign --verify --deep --strict --verbose=2 "$APP"
SIGNATURE=$(codesign -d --verbose=4 "$APP" 2>&1)
[[ "$SIGNATURE" == *'Signature=adhoc'* ]] || fail '产物不是预期的本机 ad hoc 签名。'
printf '已完成完整 ad hoc 签名校验；未公证、未安装，也未生成 DMG。\n' >&2
printf '%s\n' "$APP"
