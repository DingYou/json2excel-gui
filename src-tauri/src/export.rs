use rust_xlsxwriter::{Format, Workbook};
use serde_json::{Map, Value};
use std::collections::{BTreeSet, HashSet};

#[derive(serde::Serialize, Clone, Debug)]
pub struct ArrayFieldInfo {
    /// JSON Pointer, prefixed with root for display; escapes ~ and / unambiguously.
    pub path: String,
    pub is_matrix: bool,
    pub count: usize,
}

struct Dataset<'a> {
    path: String,
    name: String,
    rows: &'a [Value],
    headers: Option<&'a [Value]>,
}

fn child_path(path: &str, key: &str) -> String {
    format!("{}/{}", path, key.replace('~', "~0").replace('/', "~1"))
}

fn discover<'a>(value: &'a Value, path: &str, name: &str, out: &mut Vec<Dataset<'a>>) {
    match value {
        Value::Array(rows) => out.push(Dataset {
            path: path.into(),
            name: name.into(),
            rows,
            headers: None,
        }),
        Value::Object(map) => {
            // The keys array is schema metadata. Do not suppress other sibling datasets.
            let matrix = matches!((map.get("keys"), map.get("values")),
                (Some(Value::Array(_)), Some(Value::Array(rows))) if rows.iter().all(Value::is_array));
            let mut entries: Vec<_> = map.iter().collect();
            entries.sort_by_key(|(key, _)| *key);
            for (key, value) in entries {
                if matrix && key == "keys" {
                    continue;
                }
                let next = child_path(path, key);
                if matrix && key == "values" {
                    out.push(Dataset {
                        path: next,
                        name: key.clone(),
                        rows: value.as_array().expect("matrix rows"),
                        headers: map["keys"].as_array().map(Vec::as_slice),
                    });
                } else {
                    discover(value, &next, key, out);
                }
            }
        }
        _ => {}
    }
}

fn parse(content: &str) -> Result<Value, String> {
    if content.len() > 64 * 1024 * 1024 {
        return Err("JSON 内容超过 64 MiB 限制".into());
    }
    let value: Value = serde_json::from_str(content).map_err(|e| format!("解析 JSON 失败: {e}"))?;
    if !value.is_array() && !value.is_object() {
        return Err("不支持的 JSON 顶层结构（必须是对象或数组）".into());
    }
    Ok(value)
}

pub fn fields(content: &str) -> Result<Vec<ArrayFieldInfo>, String> {
    let value = parse(content)?;
    let mut datasets = Vec::new();
    discover(&value, "root", "root", &mut datasets);
    Ok(datasets
        .into_iter()
        .map(|d| ArrayFieldInfo {
            path: d.path,
            is_matrix: d.headers.is_some(),
            count: d.rows.len(),
        })
        .collect())
}

pub fn workbook(content: &str, selected: &[String]) -> Result<Workbook, String> {
    if selected.is_empty() {
        return Err("未选择任何要导出的字段".into());
    }
    let value = parse(content)?;
    let mut datasets = Vec::new();
    discover(&value, "root", "root", &mut datasets);
    let wanted: HashSet<_> = selected.iter().collect();
    if wanted.len() != selected.len() {
        return Err("选择的字段重复，请重新解析 JSON".into());
    }
    let datasets: Vec<_> = datasets
        .into_iter()
        .filter(|d| wanted.contains(&d.path))
        .collect();
    if datasets.len() != selected.len() {
        return Err("部分选中字段已不存在，请重新解析 JSON".into());
    }

    let mut book = Workbook::new();
    let mut names = HashSet::new();
    let bold = Format::new().set_bold();
    let text_format = Format::new().set_num_format("@");
    for dataset in datasets {
        let headers =
            if let Some(headers) = dataset.headers {
                let headers: Vec<_> = headers.iter().map(cell_text).collect();
                if dataset
                    .rows
                    .iter()
                    .any(|r| r.as_array().is_none_or(|r| r.len() != headers.len()))
                {
                    return Err(format!("{}: 矩阵每行列数必须与 keys 相同", dataset.path));
                }
                headers
            } else {
                let mut keys = BTreeSet::new();
                for (row, value) in dataset.rows.iter().enumerate() {
                    let record = value.as_object().ok_or_else(|| format!(
                    "{}: 第 {} 行不是对象；仅支持对象数组或 keys/values 矩阵，未导出任何文件",
                    dataset.path, row + 1))?;
                    keys.extend(record.keys().cloned());
                }
                keys.into_iter().collect()
            };
        if headers.len() > 16_384 || dataset.rows.len() > 1_048_575 {
            return Err(format!("{}: 超过 Excel 行列上限", dataset.path));
        }
        // Empty arrays intentionally produce a named blank sheet, not an implicit Sheet1.
        // A nonempty array of empty objects has no representable columns: reject, never drop rows.
        if headers.is_empty() && !dataset.rows.is_empty() {
            return Err(format!("{}: 数据行没有任何列，无法无损导出", dataset.path));
        }
        let name = unique_sheet_name(&dataset.name, &mut names);
        let sheet = book.add_worksheet();
        sheet.set_name(&name).map_err(|e| e.to_string())?;
        for (col, header) in headers.iter().enumerate() {
            sheet
                .write_string_with_format(0, col as u16, header, &bold)
                .map_err(|e| e.to_string())?;
        }
        for (row, value) in dataset.rows.iter().enumerate() {
            // A stored blank cell retains even trailing empty records/empty-string rows.
            sheet
                .write_blank((row + 1) as u32, 0, &text_format)
                .map_err(|e| e.to_string())?;
            for (col, header) in headers.iter().enumerate() {
                let cell = if dataset.headers.is_some() {
                    &value.as_array().expect("validated matrix")[col]
                } else {
                    let record: &Map<String, Value> = value.as_object().expect("validated record");
                    let Some(cell) = record.get(header) else {
                        continue;
                    };
                    cell
                };
                // All values are text, so Excel cannot round long IDs or execute formulas.
                sheet
                    .write_string((row + 1) as u32, col as u16, cell_text(cell))
                    .map_err(|e| e.to_string())?;
            }
        }
    }
    Ok(book)
}

fn cell_text(value: &Value) -> String {
    match value {
        Value::Null => "<nil>".into(),
        Value::String(s) => s.clone(),
        _ => value.to_string(),
    }
}

fn unique_sheet_name(raw: &str, used: &mut HashSet<String>) -> String {
    let base: String = raw
        .chars()
        .map(|c| {
            if c.is_control() || "[]:*?/\\".contains(c) {
                '_'
            } else {
                c
            }
        })
        .collect();
    let base = base.trim_matches(|c: char| c.is_whitespace() || c == char::from(39));

    let base = if base.is_empty() { "Sheet" } else { base };
    for index in 1.. {
        let suffix = if index == 1 {
            String::new()
        } else {
            format!(" ({index})")
        };
        // Conservatively count UTF-16 units as well as preserving scalar boundaries:
        // supplementary characters must never be byte-sliced or split into surrogates.
        let mut units = 0;
        let prefix: String = base
            .chars()
            .take_while(|c| {
                units += c.len_utf16();
                units <= 31 - suffix.len()
            })
            .collect();
        let candidate = format!("{}{}", prefix.trim_end_matches('\''), suffix);
        let folded = candidate.to_lowercase();
        // "History" is reserved by Excel's change tracking feature.
        if folded != "history" && used.insert(folded) {
            return candidate;
        }
    }
    unreachable!()
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::{Data, Reader, Xlsx};
    use std::io::Cursor;

    fn read(content: &str, selected: &[&str]) -> Xlsx<Cursor<Vec<u8>>> {
        let bytes = workbook(
            content,
            &selected.iter().map(|s| s.to_string()).collect::<Vec<_>>(),
        )
        .unwrap()
        .save_to_buffer()
        .unwrap();
        Xlsx::new(Cursor::new(bytes)).unwrap()
    }

    #[test]
    fn trailing_empty_records_are_retained_in_sheet_dimensions() {
        for (content, selected, sheet) in [
            (r#"[{"a":"one"},{"a":""},{}]"#, "root", "root"),
            (
                r#"{"keys":["a"],"values":[["one"],[""]]}"#,
                "root/values",
                "values",
            ),
        ] {
            let mut book = read(content, &[selected]);
            let mut cells = book.worksheet_cells_reader(sheet).unwrap();
            let last = if sheet == "root" { 3 } else { 2 };
            assert_eq!(cells.dimensions().end.0, last);
            let mut positions = Vec::new();
            while let Some(cell) = cells.next_cell().unwrap() {
                positions.push(cell.get_position());
            }
            assert!(positions.contains(&(last, 0)));
        }
    }

    #[test]
    fn union_headers_and_long_numbers_roundtrip() {
        let mut book = read(
            r#"[{"id":9007199254740993,"z":1},{"a":"second","id":123456789012345678901234567890},{"id":"=1+1","nested":{"n":999999999999999999999}}]"#,
            &["root"],
        );
        let range = book.worksheet_range("root").unwrap();
        assert_eq!(range.height(), 4);
        assert_eq!(
            range.rows().next().unwrap(),
            &[
                Data::String("a".into()),
                Data::String("id".into()),
                Data::String("nested".into()),
                Data::String("z".into())
            ]
        );
        assert_eq!(range[(1, 1)], Data::String("9007199254740993".into()));
        assert_eq!(
            range[(2, 1)],
            Data::String("123456789012345678901234567890".into())
        );
        assert_eq!(range[(2, 0)], Data::String("second".into()));
        assert_eq!(range[(3, 1)], Data::String("=1+1".into()));
        assert_eq!(
            range[(3, 2)],
            Data::String(r#"{"n":999999999999999999999}"#.into())
        );
    }

    #[test]
    fn matrix_siblings_multiple_selection_and_empty() {
        let content = r#"{"keys":["姓名","ID"],"values":[["甲",9007199254740993],["乙",2]],"users":[{"name":"Ada"}],"empty":[]}"#;
        assert_eq!(
            fields(content)
                .unwrap()
                .iter()
                .map(|f| f.path.as_str())
                .collect::<Vec<_>>(),
            vec!["root/empty", "root/users", "root/values"]
        );
        let mut book = read(content, &["root/values", "root/users", "root/empty"]);
        assert_eq!(book.sheet_names(), vec!["empty", "users", "values"]);
        assert!(book.worksheet_range("empty").unwrap().is_empty());
        let range = book.worksheet_range("values").unwrap();
        assert_eq!(range[(0, 0)], Data::String("姓名".into()));
        assert_eq!(range[(1, 1)], Data::String("9007199254740993".into()));
        let book = read(content, &["root/users"]);
        assert_eq!(book.sheet_names(), vec!["users"]);
        assert!(
            read("[]", &["root"])
                .worksheet_range("root")
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn unsupported_rows_fail_instead_of_disappearing() {
        for content in [
            "[1,2]",
            "[{\"a\":1},null]",
            "[[1]]",
            "[{}]",
            "{\"keys\":[\"a\"],\"values\":[[1,2]]}",
        ] {
            let selected = fields(content)
                .unwrap()
                .into_iter()
                .map(|f| f.path)
                .collect::<Vec<_>>();
            assert!(workbook(content, &selected).is_err(), "{content}");
        }
        assert!(fields("true").is_err());
        assert!(workbook("[]", &[]).is_err());
        assert!(workbook("[]", &["missing".into()]).is_err());
        assert!(workbook("[]", &["root".into(), "root".into()]).is_err());
    }

    #[test]
    fn unicode_names_and_selectors_roundtrip() {
        let content = serde_json::json!({
            "甲".repeat(40): [], "甲".repeat(39) + "乙": [],
            "A": [{"x":1}], "a": [{"x":2}], "History": [],
            "bad[]:*?\name": [], "''": [], " 'name' ": [], " ' ": [],
            "a.b": [{"x":3}], "a/b": [], "a~b": [],
            "nested": {"a.b": [{"x":4}]}, "😀".repeat(35): [],
        })
        .to_string();
        let selected: Vec<_> = fields(&content)
            .unwrap()
            .into_iter()
            .map(|f| f.path)
            .collect();
        assert!(selected.contains(&"root/a~1b".to_string()));
        assert!(selected.contains(&"root/a~0b".to_string()));
        assert!(selected.contains(&"root/a.b".to_string()));
        assert!(selected.contains(&"root/nested/a.b".to_string()));
        let bytes = workbook(&content, &selected)
            .unwrap()
            .save_to_buffer()
            .unwrap();
        let book = Xlsx::new(Cursor::new(bytes)).unwrap();
        let names = book.sheet_names();
        assert_eq!(names.len(), selected.len());
        let folded: HashSet<_> = names.iter().map(|n| n.to_lowercase()).collect();
        assert_eq!(folded.len(), names.len());
        assert!(
            names
                .iter()
                .all(|n| n.encode_utf16().count() <= 31 && !n.is_empty())
        );
        assert!(!names.iter().any(|n| n.eq_ignore_ascii_case("history")));
    }
}
