use rust_xlsxwriter::{Format, Workbook};
use serde_json::Value;
use std::collections::HashMap;
use std::fs;
use std::process::Command;
use tauri_plugin_opener::OpenerExt;

// --- DTOs ---

#[derive(serde::Serialize, Clone)]
pub struct ArrayFieldInfo {
    pub path: String,
    pub is_matrix: bool,
    pub count: usize,
}

struct ParsedField {
    path: String,
    array: Vec<Value>,
    headers: Vec<String>,
    is_matrix: bool,
}

// --- Tauri Commands ---

#[tauri::command]
async fn get_json_fields(input: String) -> Result<Vec<ArrayFieldInfo>, String> {
    let input = expand_tilde(&input);
    let content = fs::read_to_string(&input).map_err(|e| format!("读取文件失败: {}", e))?;
    parse_json_fields(&content)
}

/// 从 JSON 内容字符串获取字段信息（供 curl/粘贴模式使用）
#[tauri::command]
async fn get_json_fields_from_content(content: String) -> Result<Vec<ArrayFieldInfo>, String> {
    parse_json_fields(&content)
}

/// 将 ~ 开头的路径扩展为完整路径
fn expand_tilde(path: &str) -> String {
    if !path.starts_with('~') {
        return path.to_string();
    }

    let home = if cfg!(windows) {
        std::env::var("USERPROFILE").or_else(|_| std::env::var("HOME"))
    } else {
        std::env::var("HOME")
    };

    match home {
        Ok(home_path) => {
            if path == "~" {
                home_path
            } else if path.starts_with("~/") || path.starts_with("~\\") {
                format!("{}{}", home_path, &path[1..])
            } else {
                // ~username 格式暂时不支持，直接返回原路径
                path.to_string()
            }
        }
        Err(_) => path.to_string(),
    }
}

/// 解析 JSON 字符串并返回所有包含数组的路径
fn parse_json_fields(content: &str) -> Result<Vec<ArrayFieldInfo>, String> {
    let parsed: Value =
        serde_json::from_str(content).map_err(|e| format!("解析 JSON 失败: {}", e))?;

    let mut fields = Vec::new();

    match parsed {
        Value::Array(arr) => {
            fields.push(ArrayFieldInfo {
                path: "root".to_string(),
                is_matrix: false,
                count: arr.len(),
            });
        }
        Value::Object(map) => {
            let mut parsed_fields = Vec::new();
            find_array_fields(
                &map.into_iter().collect::<HashMap<_, _>>(),
                "root",
                false,
                &mut parsed_fields,
            );
            for f in parsed_fields {
                fields.push(ArrayFieldInfo {
                    path: f.path,
                    is_matrix: f.is_matrix,
                    count: f.array.len(),
                });
            }
        }
        _ => return Err("不支持的 JSON 顶层结构（必须是对象或数组）".to_string()),
    }

    Ok(fields)
}

/// 导出选中的字段到 Excel（文件模式）
#[tauri::command]
async fn export_to_excel(
    input: String,
    output: String,
    selected_paths: Vec<String>,
) -> Result<String, String> {
    let input = expand_tilde(&input); // Call expand_tilde
    let output = expand_tilde(&output); // Call expand_tilde
    let content = fs::read_to_string(&input).map_err(|e| format!("读取文件失败: {}", e))?;
    parse_and_export(&content, &output, selected_paths)
}

/// 导出选中的字段到 Excel（内容模式，供 curl/粘贴使用）
#[tauri::command]
async fn export_to_excel_from_content(
    content: String,
    output: String,
    selected_paths: Vec<String>,
) -> Result<String, String> {
    let output = expand_tilde(&output);
    parse_and_export(&content, &output, selected_paths)
}

/// 执行 curl 命令并返回 JSON 响应内容
#[tauri::command]
async fn execute_curl(command: String) -> Result<String, String> {
    let trimmed = command.trim();

    // 提取实际的 curl 命令参数
    let curl_args = parse_curl_command(trimmed)?;

    let output = Command::new("curl")
        .args(&curl_args)
        .arg("-s") // 静默模式，不输出进度信息
        .arg("-S") // 但仍然显示错误信息
        .output()
        .map_err(|e| format!("执行 curl 失败: {}", e))?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        let stdout = String::from_utf8_lossy(&output.stdout);
        let detail = if !stderr.is_empty() {
            stderr.to_string()
        } else if !stdout.is_empty() {
            format!("响应内容: {}", &stdout[..stdout.len().min(200)])
        } else {
            format!("退出码: {}", output.status)
        };
        return Err(format!("curl 执行出错: {}", detail));
    }

    let body = String::from_utf8(output.stdout).map_err(|e| format!("解析响应内容失败: {}", e))?;

    // 验证返回的是有效 JSON
    serde_json::from_str::<Value>(&body).map_err(|e| format!("返回内容不是有效 JSON: {}", e))?;

    Ok(body)
}

/// 从 JSON 内容解析并导出 Excel
fn parse_and_export(
    content: &str,
    output: &str,
    selected_paths: Vec<String>,
) -> Result<String, String> {
    if selected_paths.is_empty() {
        return Err("未选择任何要导出的字段".to_string());
    }

    let parsed: Value =
        serde_json::from_str(content).map_err(|e| format!("解析 JSON 失败: {}", e))?;

    let mut all_fields = Vec::new();
    match parsed {
        Value::Array(arr) => {
            all_fields.push(ParsedField {
                path: "root".to_string(),
                array: arr,
                headers: vec![],
                is_matrix: false,
            });
        }
        Value::Object(map) => {
            find_array_fields(
                &map.into_iter().collect::<HashMap<_, _>>(),
                "root",
                false,
                &mut all_fields,
            );
        }
        _ => return Err("不支持的 JSON 顶层结构".to_string()),
    }

    let datasets: Vec<ParsedField> = all_fields
        .into_iter()
        .filter(|f| selected_paths.contains(&f.path))
        .collect();

    if datasets.is_empty() {
        return Err("未找到选中的字段数据".to_string());
    }

    write_excel_sheets(datasets, output).map_err(|e| format!("导出 Excel 失败: {}", e))?;

    Ok("转换成功！".to_string())
}

/// 解析用户粘贴的 curl 命令，提取参数列表
fn parse_curl_command(input: &str) -> Result<Vec<String>, String> {
    let trimmed = input.trim();

    // 支持 fetch 格式：简单提取 URL 并转为 curl
    if trimmed.starts_with("fetch(") || trimmed.starts_with("fetch (") {
        return parse_fetch_to_curl_args(trimmed);
    }

    // 去掉开头的 curl 命令名
    let args_str = if trimmed.to_lowercase().starts_with("curl ") {
        &trimmed[5..]
    } else {
        // 如果不以 curl 开头，假设用户只粘贴了参数
        trimmed
    };

    // 简单的 shell 参数解析（处理引号）
    shell_split(args_str)
}

/// 将 fetch 格式转换为 curl 参数
fn parse_fetch_to_curl_args(input: &str) -> Result<Vec<String>, String> {
    // 提取 fetch 中的 URL（第一个引号对中的内容）
    let url =
        extract_quoted_string(input).ok_or_else(|| "无法从 fetch 命令中提取 URL".to_string())?;

    let mut args = vec![url];

    // 提取选项对象：找到 URL 后的第一个 { 到匹配的 }
    if let Some(options_json) = extract_fetch_options(input) {
        // 解析为 JSON 对象
        match serde_json::from_str::<Value>(&options_json) {
            Ok(Value::Object(opts)) => {
                // 提取 headers
                if let Some(Value::Object(headers)) = opts.get("headers") {
                    for (key, value) in headers {
                        if let Value::String(val) = value {
                            args.push("-H".to_string());
                            args.push(format!("{}: {}", key, val));
                        }
                    }
                }

                // 提取 method
                if let Some(Value::String(method)) = opts.get("method")
                    && method.to_uppercase() != "GET" {
                        args.push("-X".to_string());
                        args.push(method.to_uppercase());
                    }

                // 提取 body
                if let Some(Value::String(body)) = opts.get("body") {
                    args.push("-d".to_string());
                    args.push(body.clone());
                }

                // 提取 credentials（include -> 发送 cookies）
                // curl 默认就会发送 cookies，所以不需要额外参数
            }
            Ok(_) => {
                // 选项不是对象，忽略
            }
            Err(e) => {
                return Err(format!("解析 fetch 选项失败: {}", e));
            }
        }
    }

    Ok(args)
}

/// 从 fetch 调用中提取选项 JSON 对象
fn extract_fetch_options(input: &str) -> Option<String> {
    // 找到 URL 之后的第一个 ','，然后提取后面的 {...}
    let chars: Vec<char> = input.chars().collect();
    let mut i = 0;

    // 跳过 "fetch" 和 "("
    while i < chars.len() && chars[i] != '(' {
        i += 1;
    }
    if i >= chars.len() {
        return None;
    }
    i += 1; // 跳过 '('

    // 跳过 URL 字符串（找到匹配的引号对）
    while i < chars.len() && chars[i] != '"' && chars[i] != '\'' {
        i += 1;
    }
    if i >= chars.len() {
        return None;
    }
    let quote = chars[i];
    i += 1;
    while i < chars.len() && chars[i] != quote {
        if chars[i] == '\\' {
            i += 1; // 跳过转义字符
        }
        i += 1;
    }
    i += 1; // 跳过结束引号

    // 找到 ',' 分隔符
    while i < chars.len() && chars[i] != ',' {
        i += 1;
    }
    if i >= chars.len() {
        return None; // 没有选项对象
    }
    i += 1; // 跳过 ','

    // 跳过空白
    while i < chars.len() && chars[i].is_whitespace() {
        i += 1;
    }

    // 找到 '{' 开始
    if i >= chars.len() || chars[i] != '{' {
        return None;
    }

    // 匹配大括号，找到对应的 '}'
    let start = i;
    let mut depth = 0;
    let mut in_str = false;
    let mut str_char = '"';
    let mut escape = false;

    while i < chars.len() {
        if escape {
            escape = false;
            i += 1;
            continue;
        }
        match chars[i] {
            '\\' if in_str => {
                escape = true;
            }
            c if c == str_char && in_str => {
                in_str = false;
            }
            '"' | '\'' if !in_str => {
                in_str = true;
                str_char = chars[i];
            }
            '{' if !in_str => {
                depth += 1;
            }
            '}' if !in_str => {
                depth -= 1;
                if depth == 0 {
                    return Some(chars[start..=i].iter().collect());
                }
            }
            _ => {}
        }
        i += 1;
    }

    None
}

/// 从字符串中提取第一个引号包围的内容
fn extract_quoted_string(s: &str) -> Option<String> {
    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '"' || chars[i] == '\'' {
            let quote = chars[i];
            i += 1;
            let start = i;
            while i < chars.len() && chars[i] != quote {
                i += 1;
            }
            if i < chars.len() {
                return Some(chars[start..i].iter().collect());
            }
        }
        i += 1;
    }
    None
}

/// 简单的 shell 参数分割（处理单引号和双引号）
fn shell_split(input: &str) -> Result<Vec<String>, String> {
    let mut args = Vec::new();
    let mut current = String::new();
    let mut in_single_quote = false;
    let mut in_double_quote = false;
    let mut escape_next = false;

    for ch in input.chars() {
        if escape_next {
            // \ + 换行 -> 续行符，忽略两者
            if ch == '\n' || ch == '\r' {
                escape_next = false;
                continue;
            }
            current.push(ch);
            escape_next = false;
            continue;
        }

        match ch {
            '\\' if !in_single_quote => {
                escape_next = true;
            }
            '\'' if !in_double_quote => {
                in_single_quote = !in_single_quote;
            }
            '"' if !in_single_quote => {
                in_double_quote = !in_double_quote;
            }
            ' ' | '\t' | '\n' if !in_single_quote && !in_double_quote => {
                if !current.is_empty() {
                    args.push(current.clone());
                    current.clear();
                }
            }
            _ => {
                current.push(ch);
            }
        }
    }

    if !current.is_empty() {
        args.push(current);
    }

    if in_single_quote || in_double_quote {
        return Err("命令中存在未闭合的引号".to_string());
    }

    Ok(args)
}

/// 打开系统文件管理器并高亮选中的文件
#[tauri::command]
fn reveal_in_explorer(app: tauri::AppHandle, path: String) -> Result<(), String> {
    let path = expand_tilde(&path);
    app.opener()
        .reveal_item_in_dir(&path)
        .map_err(|e| e.to_string())
}

// --- 核心逻辑 ---

fn find_array_fields(
    m: &HashMap<String, Value>,
    path: &str,
    inside_matrix: bool,
    results: &mut Vec<ParsedField>,
) {
    // 检查是否有 keys+values 二维表结构
    if !inside_matrix
        && let (Some(Value::Array(keys_raw)), Some(Value::Array(values))) =
            (m.get("keys"), m.get("values"))
        {
            let is_matrix = values.iter().all(|row| row.is_array());
            if is_matrix {
                let headers: Vec<String> = keys_raw
                    .iter()
                    .map(|h| match h {
                        Value::String(s) => s.clone(),
                        other => other.to_string(),
                    })
                    .collect();

                results.push(ParsedField {
                    path: format!("{}.values", path),
                    array: values.clone(),
                    headers,
                    is_matrix: true,
                });
                return; // 已经是矩阵了，不再往内部递归
            }
        }

    // 保证顺序输出，先收集所有的 key 并排序
    let mut sorted_keys: Vec<&String> = m.keys().collect();
    sorted_keys.sort();

    for k in sorted_keys {
        let v = &m[k];
        let new_path = format!("{}.{}", path, k);
        match v {
            Value::Array(arr) => {
                results.push(ParsedField {
                    path: new_path.clone(),
                    array: arr.clone(),
                    headers: vec![],
                    is_matrix: false,
                });
            }
            Value::Object(sub_map) => {
                find_array_fields(
                    &sub_map.clone().into_iter().collect::<HashMap<_, _>>(),
                    &new_path,
                    inside_matrix,
                    results,
                );
            }
            _ => {}
        }
    }
}

fn write_excel_sheets(datasets: Vec<ParsedField>, output_path: &str) -> Result<(), String> {
    let mut workbook = Workbook::new();
    let header_format = Format::new().set_bold();

    for set in datasets {
        if set.array.is_empty() {
            continue;
        }

        let sheet_name = clean_sheet_name(&set.path);
        let worksheet = workbook
            .add_worksheet()
            .set_name(&sheet_name)
            .map_err(|e| e.to_string())?;

        let mut row_idx = 0;

        if set.is_matrix {
            // 写入表头
            for (col_idx, h) in set.headers.iter().enumerate() {
                worksheet
                    .write_string_with_format(row_idx, col_idx as u16, h, &header_format)
                    .map_err(|e| e.to_string())?;
            }
            row_idx += 1;

            // 写入矩阵数据
            for row_val in set.array {
                if let Value::Array(row_arr) = row_val {
                    for (col_idx, cell_val) in row_arr.iter().enumerate() {
                        let cell_str = val_to_string(cell_val);
                        worksheet
                            .write_string(row_idx, col_idx as u16, &cell_str)
                            .map_err(|e| e.to_string())?;
                    }
                    row_idx += 1;
                }
            }
            continue;
        }

        // 常规对象数组
        let first_row = match &set.array[0] {
            Value::Object(o) => o,
            _ => continue, // 不是对象数组，跳过
        };

        // 收集表头（使用 BTreeMap 保持键的字母顺序或者保持原序，Rust 中 serde_json::Map 保持插入顺序）
        let headers: Vec<String> = first_row.keys().cloned().collect();
        // Option: sorted keys
        // headers.sort();

        // 写入表头
        for (col_idx, h) in headers.iter().enumerate() {
            worksheet
                .write_string_with_format(row_idx, col_idx as u16, h, &header_format)
                .map_err(|e| e.to_string())?;
        }
        row_idx += 1;

        // 写入数据
        for row_val in set.array {
            if let Value::Object(rec) = row_val {
                for (col_idx, key) in headers.iter().enumerate() {
                    if let Some(val) = rec.get(key) {
                        let cell_str = match val {
                            Value::Object(_) | Value::Array(_) => val.to_string(), // 序列化为 JSON
                            _ => val_to_string(val),
                        };
                        worksheet
                            .write_string(row_idx, col_idx as u16, &cell_str)
                            .map_err(|e| e.to_string())?;
                    }
                }
                row_idx += 1;
            }
        }
    }

    workbook.save(output_path).map_err(|e| e.to_string())?;
    Ok(())
}

fn val_to_string(val: &Value) -> String {
    match val {
        Value::Null => "<nil>".to_string(),
        Value::Bool(b) => b.to_string(),
        Value::Number(n) => n.to_string(),
        Value::String(s) => s.clone(),
        Value::Array(_) | Value::Object(_) => val.to_string(),
    }
}

fn clean_sheet_name(path: &str) -> String {
    // 提取最后一段
    let base = path.split('.').next_back().unwrap_or(path);
    // 用 _ 替换 .
    let name = base.replace('.', "_");
    // Excel 限制 sheet 名最长 31 字符
    if name.len() > 31 {
        name[..31].to_string()
    } else {
        name
    }
}

// --- Tauri 入口 ---

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_window_state::Builder::default().build())
        .invoke_handler(tauri::generate_handler![
            get_json_fields,
            get_json_fields_from_content,
            export_to_excel,
            export_to_excel_from_content,
            execute_curl,
            reveal_in_explorer
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
