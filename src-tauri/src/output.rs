use crate::export;
use same_file::Handle;
use sha2::{Digest, Sha256};
use std::fs::{self, File};
use std::io::Read;
use std::path::{Path, PathBuf};
use std::time::SystemTime;

const MAX_JSON_BYTES: u64 = 64 * 1024 * 1024;

pub fn expand_tilde(path: &str) -> PathBuf {
    if (path == "~" || path.starts_with("~/") || path.starts_with("~\\"))
        && let Some(home) = std::env::var_os(if cfg!(windows) { "USERPROFILE" } else { "HOME" })
    {
        return Path::new(&home).join(path.get(2..).unwrap_or(""));
    }
    PathBuf::from(path)
}

pub fn read_source(path: &Path) -> Result<(String, Handle), String> {
    let file = File::open(path).map_err(|e| format!("读取文件失败: {e}"))?;
    let handle = Handle::from_file(file.try_clone().map_err(|e| e.to_string())?)
        .map_err(|e| e.to_string())?;
    let mut content = String::new();
    file.take(MAX_JSON_BYTES + 1)
        .read_to_string(&mut content)
        .map_err(|e| format!("读取文件失败: {e}"))?;
    if content.len() as u64 > MAX_JSON_BYTES {
        return Err("JSON 文件超过 64 MiB 限制".into());
    }
    Ok((content, handle))
}

struct Snapshot {
    handle: Handle,
    modified: SystemTime,
    len: u64,
    digest: [u8; 32],
}

fn snapshot(path: &Path, source: Option<&Handle>) -> Result<Option<Snapshot>, String> {
    let metadata = match fs::symlink_metadata(path) {
        Ok(m) => m,
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(e) => return Err(format!("无法检查输出文件: {e}")),
    };
    if metadata.file_type().is_symlink() {
        return Err("输出文件不能是符号链接，请选择新的 .xlsx 文件".into());
    }
    if !metadata.is_file() {
        return Err("输出位置不是普通文件".into());
    }
    let mut file = File::open(path).map_err(|e| format!("无法检查输出文件: {e}"))?;
    let metadata = file.metadata().map_err(|e| e.to_string())?;
    let mut hash = Sha256::new();
    let mut read = 0u64;
    let mut buffer = [0; 65_536];
    loop {
        let count = file
            .read(&mut buffer)
            .map_err(|e| format!("读取输出文件失败: {e}"))?;
        if count == 0 {
            break;
        }
        read += count as u64;
        if read > metadata.len() {
            return Err("输出文件正在变化，请稍后重试".into());
        }
        hash.update(&buffer[..count]);
    }
    let handle = Handle::from_file(file).map_err(|e| format!("无法检查输出文件: {e}"))?;
    if source.is_some_and(|source| source == &handle) {
        return Err("输出文件不能是源 JSON 文件或其硬链接/路径别名".into());
    }
    if read != metadata.len() {
        return Err("输出文件正在变化，请稍后重试".into());
    }
    Ok(Some(Snapshot {
        handle,
        modified: metadata.modified().map_err(|e| e.to_string())?,
        len: metadata.len(),
        digest: hash.finalize().into(),
    }))
}

fn destination(output: &Path) -> Result<PathBuf, String> {
    if !output
        .extension()
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx"))
    {
        return Err("输出文件必须使用 .xlsx 扩展名，不能覆盖 JSON 文件".into());
    }
    let name = output.file_name().ok_or("输出文件名无效")?;
    let parent = output
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let parent = parent
        .canonicalize()
        .map_err(|e| format!("输出目录不存在或不可访问: {e}"))?;
    Ok(parent.join(name))
}

/// Stages a complete workbook in the same directory before asking to replace or publishing.
/// Confirmation is owned by the backend; invoke callers cannot bypass it with a boolean flag.
pub fn export(
    content: &str,
    output: &Path,
    selected: &[String],
    source: Option<&Handle>,
    confirm_replace: impl FnOnce(&Path) -> bool,
) -> Result<String, String> {
    let output = destination(output)?;
    let before = snapshot(&output, source)?;
    let mut workbook = export::workbook(content, selected)?;
    let mut staged = tempfile::Builder::new()
        .prefix(".json2excel-")
        .suffix(".tmp")
        .tempfile_in(output.parent().expect("canonical parent"))
        .map_err(|e| format!("创建临时文件失败: {e}"))?;
    workbook
        .save_to_writer(staged.as_file_mut())
        .map_err(|e| format!("导出 Excel 失败: {e}"))?;
    staged
        .as_file()
        .sync_all()
        .map_err(|e| format!("写入临时文件失败: {e}"))?;

    if before.is_some() && !confirm_replace(&output) {
        return Err("已取消覆盖，原文件保持不变".into());
    }
    let after = snapshot(&output, source)?;
    match (&before, &after) {
        (None, None) => {
            staged
                .persist_noclobber(&output)
                .map_err(|e| format!("输出位置已被占用或写入失败，未覆盖: {}", e.error))?;
        }
        (Some(before), Some(after))
            if before.handle == after.handle
                && before.modified == after.modified
                && before.len == after.len
                && before.digest == after.digest =>
        {
            // Atomic replacement, never open/truncate the destination (or any hardlink to it).
            staged
                .persist(&output)
                .map_err(|e| format!("替换输出文件失败，原文件保持不变: {}", e.error))?;
        }
        _ => return Err("输出文件在转换期间发生变化，未覆盖；请重新转换并确认".into()),
    }
    Ok("转换成功！".into())
}

#[cfg(test)]
mod tests {
    use super::*;
    fn selected() -> Vec<String> {
        vec!["root".into()]
    }
    const JSON: &str = r#"[{"id":9007199254740993}]"#;

    #[test]
    fn same_size_and_timestamp_edits_still_invalidate_confirmation() {
        let dir = tempfile::tempdir().unwrap();
        let output = dir.path().join("out.xlsx");
        fs::write(&output, "original").unwrap();
        let modified = fs::metadata(&output).unwrap().modified().unwrap();
        assert!(
            export(JSON, &output, &selected(), None, |path| {
                fs::write(path, "modified").unwrap();
                File::options()
                    .write(true)
                    .open(path)
                    .unwrap()
                    .set_modified(modified)
                    .unwrap();
                true
            })
            .is_err()
        );
        assert_eq!(fs::read_to_string(&output).unwrap(), "modified");
    }

    #[test]
    fn new_existing_cancel_confirm_and_validation_failure() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("result.xlsx");
        export(JSON, &path, &selected(), None, |_| {
            panic!("new output needs no prompt")
        })
        .unwrap();
        let original = fs::read(&path).unwrap();
        assert!(export(JSON, &path, &selected(), None, |_| false).is_err());
        assert_eq!(fs::read(&path).unwrap(), original);
        assert!(
            export("[1]", &path, &selected(), None, |_| panic!(
                "invalid workbook"
            ))
            .is_err()
        );
        assert_eq!(fs::read(&path).unwrap(), original);
        export("[{\"new\":true}]", &path, &selected(), None, |_| true).unwrap();
        assert_ne!(fs::read(&path).unwrap(), original);
        assert_eq!(
            fs::read_dir(dir.path()).unwrap().count(),
            1,
            "staged files cleaned"
        );
    }

    #[test]
    fn refuses_source_aliases_and_non_xlsx_outputs() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xlsx");
        fs::write(&source, JSON).unwrap();
        let (content, handle) = read_source(&source).unwrap();
        let hard = dir.path().join("hard.xlsx");
        fs::hard_link(&source, &hard).unwrap();
        for path in [&source, &hard, &dir.path().join("source.json")] {
            assert!(export(&content, path, &selected(), Some(&handle), |_| true).is_err());
        }
        assert_eq!(fs::read_to_string(&source).unwrap(), JSON);
        #[cfg(unix)]
        {
            use std::os::unix::fs::symlink;
            let alias = dir.path().join("alias.xlsx");
            symlink(&source, &alias).unwrap();
            assert!(export(&content, &alias, &selected(), Some(&handle), |_| true).is_err());
            let parent_alias = dir.path().join("parent");
            symlink(dir.path(), &parent_alias).unwrap();
            assert!(
                export(
                    &content,
                    &parent_alias.join("source.xlsx"),
                    &selected(),
                    Some(&handle),
                    |_| true
                )
                .is_err()
            );
        }
    }

    #[test]
    fn changed_destination_is_not_overwritten() {
        let dir = tempfile::tempdir().unwrap();
        let output = dir.path().join("out.xlsx");
        fs::write(&output, "original").unwrap();
        assert!(
            export(JSON, &output, &selected(), None, |path| {
                fs::write(path, "changed by another process").unwrap();
                true
            })
            .is_err()
        );
        assert_eq!(
            fs::read_to_string(&output).unwrap(),
            "changed by another process"
        );
    }
}
