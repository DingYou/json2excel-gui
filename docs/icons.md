# 桌面图标资源

## 源图与范围

- 规范源图：仓库根目录 `app-icon.png`，用户确认的 **A 奶白浅蓝紫版**，1254×1254、8-bit RGBA。原文件名为 `json2excel-icon-A-milky-periwinkle.png`；保留原始像素与透明通道，不重新绘制或用旧图替代。外侧 6% 边带的原始 alpha 最大为 2，保持原样，不额外作阈值清理或前景调整。
- 源图 SHA-256：`4de24f79619162a1c7e29c47908511a46de1b635efcb11de6369ae0e95032310`。
- 原根目录 SVG 已归档为 `docs/legacy/app-icon-v1.svg`，仅供历史参考，不再作为生成输入；当前构建配置未引用它。
- `src-tauri/icons/` 顶层 PNG、Windows ICO 和 macOS ICNS 均由同一源图生成。包含 Linux 通用的 32/64/128/256/512px PNG、Windows 小尺寸 ICO/Logo，以及 macOS 最高 1024px 的 ICNS 图像帧。
- 本次只更新桌面资源。现存 `src-tauri/icons/ios/` 保持原样；生成器附带的 iOS/Android 输出不复制到项目，未宣称移动端已适配。

## 重新生成

使用 Cargo Tauri CLI v2（实测 2.10.0，准备方式见根 README），从仓库根目录执行。生成和格式转换由 Tauri 原生工具完成，不需要 Node/npm：

```bash
cargo tauri --version
STAGE=$(mktemp -d "${TMPDIR:-/tmp}/json2excel-icons.XXXXXX")
cargo tauri icon app-icon.png --output "$STAGE"
# 只复制桌面文件，不覆盖任何移动端子目录。
for file in "$STAGE"/*; do
  if [ -f "$file" ]; then
    case "$file" in
      *.png|*.ico|*.icns) cp "$file" src-tauri/icons/ ;;
    esac
  fi
done
```

`tauri.conf.json` 中五个现有桌面图标路径保持不变。保留规范源图的原始 alpha，不添加底色、不裁切、不替换玻璃效果；小尺寸由同一生成器缩放。

## 验证

```bash
# Node 仅用于资源/前端回归测试，不是生成或 macOS 构建的依赖。
npm run check
# macOS 本机完整 ad hoc 签名的 .app 构建，不自动安装或生成 DMG。
bash scripts/build-macos.sh aarch64-apple-darwin
```

图标回归测试检查规范源图的 SHA-256、PNG 尺寸与实际透明 alpha、配置引用、ICO 多尺寸 32-bit PNG 帧，以及 ICNS 图像帧/小尺寸 alpha mask。若以后明确更换源图，需同步更新测试中的规范源图校验值。

本次已查看原图及 32/64/128px 实际像素，并在浅/深背景上用原生 CoreGraphics 合成检查透明边缘；另用 macOS `iconutil` 解出 ICNS 各档尺寸；macOS 本机构建及完整签名校验已通过。Windows/Linux 资源格式已检查，但未在对应系统运行安装器或检查系统图标缓存。图标处理不涉及 Developer ID、公证或系统安全设置修改。
