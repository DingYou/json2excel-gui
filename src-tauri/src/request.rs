//! Data-only request import. No shell, JavaScript interpreter, curl config, or file references.
use reqwest::{
    Method, Url,
    header::{HeaderMap, HeaderName, HeaderValue},
};
use serde_json::{Map, Value};
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::Duration;

const MAX_COMMAND: usize = 1024 * 1024;
const MAX_RESPONSE: usize = 16 * 1024 * 1024;
static REQUEST_ACTIVE: AtomicBool = AtomicBool::new(false);
struct RequestGuard;
impl Drop for RequestGuard {
    fn drop(&mut self) {
        REQUEST_ACTIVE.store(false, Ordering::Release);
    }
}

#[derive(Debug)]
struct Request {
    url: Url,
    method: Method,
    headers: HeaderMap,
    body: Option<String>,
}

pub async fn execute(command: &str) -> Result<String, String> {
    REQUEST_ACTIVE
        .compare_exchange(false, true, Ordering::Acquire, Ordering::Relaxed)
        .map_err(|_| "已有请求正在执行，请等待完成".to_string())?;
    let _guard = RequestGuard;
    let request = parse(command)?;
    send(request, Duration::from_secs(30)).await
}

async fn send(request: Request, timeout: Duration) -> Result<String, String> {
    let client = reqwest::Client::builder()
        .timeout(timeout)
        .connect_timeout(Duration::from_secs(10))
        .redirect(reqwest::redirect::Policy::none())
        .build()
        .map_err(|_| "无法初始化 HTTP 客户端".to_string())?;
    let mut builder = client
        .request(request.method, request.url)
        .headers(request.headers);
    if let Some(body) = request.body {
        builder = builder.body(body);
    }
    let mut response = builder
        .send()
        .await
        .map_err(|e| format!("HTTP 请求失败: {}", e.without_url()))?;
    if !response.status().is_success() {
        return Err(format!(
            "HTTP 状态 {}；不会自动跟随重定向",
            response.status()
        ));
    }
    if response
        .content_length()
        .is_some_and(|n| n > MAX_RESPONSE as u64)
    {
        return Err("响应超过 16 MiB 限制".into());
    }
    let mut bytes = Vec::new();
    while let Some(chunk) = response
        .chunk()
        .await
        .map_err(|e| format!("读取响应失败: {}", e.without_url()))?
    {
        if chunk.len() > MAX_RESPONSE - bytes.len() {
            return Err("响应超过 16 MiB 限制".into());
        }
        bytes.extend_from_slice(&chunk);
    }
    let body = String::from_utf8(bytes).map_err(|_| "响应不是有效 UTF-8".to_string())?;
    serde_json::from_str::<Value>(&body).map_err(|e| format!("返回内容不是有效 JSON: {e}"))?;
    Ok(body)
}

fn parse(command: &str) -> Result<Request, String> {
    if command.len() > MAX_COMMAND {
        return Err("请求命令超过 1 MiB 限制".into());
    }
    let command = command.trim();
    if command
        .strip_prefix("fetch")
        .is_some_and(|rest| rest.trim_start().starts_with('('))
    {
        parse_fetch(command)
    } else {
        parse_curl(command)
    }
}

fn make_request(
    url: &str,
    method: &str,
    raw_headers: Vec<(String, String)>,
    body: Option<String>,
) -> Result<Request, String> {
    if url.len() > 16_384 {
        return Err("URL 过长".into());
    }
    let url = Url::parse(url).map_err(|_| "URL 无效".to_string())?;
    if !matches!(url.scheme(), "http" | "https")
        || url.host_str().is_none()
        || !url.username().is_empty()
        || url.password().is_some()
    {
        return Err("仅支持不含用户名密码的 HTTP(S) URL".into());
    }
    let method = method.to_ascii_uppercase();
    if !matches!(
        method.as_str(),
        "GET" | "HEAD" | "POST" | "PUT" | "PATCH" | "DELETE" | "OPTIONS"
    ) {
        return Err("不支持的 HTTP 方法".into());
    }
    if raw_headers.len() > 128
        || raw_headers
            .iter()
            .map(|(k, v)| k.len() + v.len())
            .sum::<usize>()
            > 65_536
    {
        return Err("请求头超过限制（128 项 / 64 KiB）".into());
    }
    let mut headers = HeaderMap::new();
    for (key, value) in raw_headers {
        let key =
            HeaderName::from_bytes(key.as_bytes()).map_err(|_| "请求头名称无效".to_string())?;
        // Let the client negotiate only compression it can decode under the response size limit.
        if key == "accept-encoding" {
            continue;
        }
        if matches!(
            key.as_str(),
            "host"
                | "content-length"
                | "transfer-encoding"
                | "connection"
                | "upgrade"
                | "te"
                | "trailer"
        ) || key.as_str().starts_with("proxy-")
        {
            return Err(format!("不支持手动设置传输控制请求头: {key}"));
        }
        let value = HeaderValue::from_str(&value).map_err(|_| "请求头内容无效".to_string())?;
        headers.append(key, value);
    }
    Ok(Request {
        url,
        method: Method::from_bytes(method.as_bytes()).expect("allowlisted method"),
        headers,
        body,
    })
}

fn parse_fetch(input: &str) -> Result<Request, String> {
    let mut parser = DataParser::new(input);
    parser.literal("fetch")?;
    parser.punctuation('(')?;
    let url = parser.string()?;
    let options = if parser.consume(',') {
        let value = parser.value(0)?;
        value
            .as_object()
            .cloned()
            .ok_or("fetch 选项必须是数据对象")?
    } else {
        Map::new()
    };
    parser.consume(',');
    parser.punctuation(')')?;
    parser.consume(';');
    parser.whitespace();
    if parser.pos != parser.chars.len() {
        return Err("fetch 调用后不能包含其他表达式".into());
    }

    let mut method = "GET".to_string();
    let mut headers: Vec<(String, String)> = Vec::new();
    let mut body = None;
    for (key, value) in options {
        match key.as_str() {
            "method" => method = value.as_str().ok_or("method 必须是字符串")?.to_string(),
            "body" => match value {
                Value::Null => {}
                Value::String(s) => body = Some(s),
                _ => return Err("body 只支持字符串或 null".into()),
            },
            "headers" => {
                let map = value.as_object().ok_or("headers 必须是字符串键值对象")?;
                for (key, value) in map {
                    headers.push((
                        key.clone(),
                        value.as_str().ok_or("请求头值必须是字符串")?.to_string(),
                    ));
                }
            }
            // Chrome metadata is accepted as data, not browser session access or credentials.
            "mode" | "credentials" | "referrer" | "referrerPolicy" | "cache" | "priority" => {
                if !value.is_string() {
                    return Err(format!("{key} 必须是字符串"));
                }
            }
            "signal" if value.is_null() => {}
            "keepalive" if value.is_boolean() => {}
            "integrity" if value.as_str() == Some("") => {}
            "redirect" if matches!(value.as_str(), Some("error" | "manual")) => {}
            _ => return Err(format!("不支持的 fetch 选项: {key}")),
        }
    }
    if body.is_some() && matches!(method.to_ascii_uppercase().as_str(), "GET" | "HEAD") {
        return Err("fetch GET/HEAD 请求不能包含 body".into());
    }
    if body.is_some()
        && !headers
            .iter()
            .any(|(k, _)| k.eq_ignore_ascii_case("content-type"))
    {
        headers.push(("content-type".into(), "text/plain;charset=UTF-8".into()));
    }
    make_request(&url, &method, headers, body)
}

fn parse_curl(input: &str) -> Result<Request, String> {
    let args = shell_words(input)?;
    let mut index = usize::from(args.first().is_some_and(|s| s == "curl"));
    let mut url = None;
    let mut method = None;
    let mut headers: Vec<(String, String)> = Vec::new();
    let mut body = None;
    while index < args.len() {
        let arg = &args[index];
        index += 1;
        match arg.as_str() {
            "-s" | "-S" | "-sS" | "--silent" | "--show-error" | "--compressed" => {}
            "-X" | "--request" | "-H" | "--header" | "-d" | "--data" | "--data-raw"
            | "--data-binary" | "-b" | "--cookie" | "--url" => {
                let value = args.get(index).ok_or_else(|| format!("{arg} 缺少参数"))?;
                index += 1;
                match arg.as_str() {
                    "-X" | "--request" => {
                        if method.replace(value.clone()).is_some() {
                            return Err("HTTP 方法重复".into());
                        }
                    }
                    "-H" | "--header" => {
                        let (key, value) =
                            value.split_once(':').ok_or("请求头必须为 名称: 内容")?;
                        headers.push((key.trim().into(), value.trim_start().into()));
                    }
                    "-b" | "--cookie" => {
                        if !value.contains('=') || value.starts_with('@') {
                            return Err("Cookie 仅支持直接键值，不支持文件".into());
                        }
                        headers.push(("cookie".into(), value.clone()));
                    }
                    "--url" => {
                        if url.replace(value.clone()).is_some() {
                            return Err("只允许一个 URL".into());
                        }
                    }
                    _ => {
                        if arg != "--data-raw" && value.starts_with('@') {
                            return Err("请求体不支持读取本地文件".into());
                        }
                        if body.replace(value.clone()).is_some() {
                            return Err("只允许一个请求体参数".into());
                        }
                    }
                }
            }
            value if !value.starts_with('-') => {
                if url.replace(value.into()).is_some() {
                    return Err("只允许一个 URL，不支持 shell 组合命令".into());
                }
            }
            _ => return Err(format!("不支持的 curl 参数: {arg}")),
        }
    }
    if body.is_some()
        && !headers
            .iter()
            .any(|(k, _)| k.eq_ignore_ascii_case("content-type"))
    {
        headers.push((
            "content-type".into(),
            "application/x-www-form-urlencoded".into(),
        ));
    }
    let method = method.unwrap_or_else(|| {
        if body.is_some() {
            "POST".into()
        } else {
            "GET".into()
        }
    });
    make_request(&url.ok_or("缺少请求 URL")?, &method, headers, body)
}

/// POSIX-style argument decoding only. Metacharacters are never evaluated.
fn shell_words(input: &str) -> Result<Vec<String>, String> {
    let mut words = Vec::new();
    let mut word = String::new();
    let mut quote = None;
    let mut started = false;
    let mut chars = input.chars().peekable();
    while let Some(ch) = chars.next() {
        match ch {
            '\'' | '"' if quote.is_none() => {
                quote = Some(ch);
                started = true;
            }
            ch if Some(ch) == quote => {
                quote = None;
            }
            '\\' if quote != Some('\'') => {
                let next = chars.next().ok_or("命令末尾存在未完成转义")?;
                if next == '\n' {
                    continue;
                }
                if quote == Some('"') && !matches!(next, '$' | '`' | '"' | '\\') {
                    word.push('\\');
                }
                word.push(next);
                started = true;
            }
            ch if ch.is_whitespace() && quote.is_none() => {
                if started {
                    words.push(std::mem::take(&mut word));
                    started = false;
                }
            }
            ch => {
                word.push(ch);
                started = true;
            }
        }
    }
    if quote.is_some() {
        return Err("命令中存在未闭合引号".into());
    }
    if started {
        words.push(word);
    }
    Ok(words)
}

/// A deliberately small JavaScript *data literal* grammar: quoted strings, objects,
/// arrays, JSON numbers, true/false/null, identifier keys, and trailing commas.
struct DataParser {
    chars: Vec<char>,
    pos: usize,
    nodes: usize,
}
impl DataParser {
    fn new(s: &str) -> Self {
        Self {
            chars: s.chars().collect(),
            pos: 0,
            nodes: 0,
        }
    }
    fn whitespace(&mut self) {
        while self.chars.get(self.pos).is_some_and(|c| c.is_whitespace()) {
            self.pos += 1;
        }
    }
    fn consume(&mut self, c: char) -> bool {
        self.whitespace();
        if self.chars.get(self.pos) == Some(&c) {
            self.pos += 1;
            true
        } else {
            false
        }
    }
    fn punctuation(&mut self, c: char) -> Result<(), String> {
        if self.consume(c) {
            Ok(())
        } else {
            Err(format!("数据语法错误，缺少 {c}"))
        }
    }
    fn literal(&mut self, word: &str) -> Result<(), String> {
        self.whitespace();
        for ch in word.chars() {
            if self.chars.get(self.pos) != Some(&ch) {
                return Err("仅支持数据字面量，不能执行 JavaScript".into());
            }
            self.pos += 1;
        }
        Ok(())
    }
    fn string(&mut self) -> Result<String, String> {
        self.whitespace();
        let quote = *self.chars.get(self.pos).ok_or("缺少字符串")?;
        if !matches!(quote, '\'' | '"') {
            return Err("必须使用单引号或双引号字符串".into());
        }
        self.pos += 1;
        let mut json = String::from("\"");
        while let Some(&ch) = self.chars.get(self.pos) {
            self.pos += 1;
            if ch == quote {
                json.push('"');
                return serde_json::from_str(&json).map_err(|_| "无效字符串转义".into());
            }
            if ch == '\\' {
                let next = *self.chars.get(self.pos).ok_or("未完成字符串转义")?;
                self.pos += 1;
                match next {
                    '\'' => json.push('\''),
                    'x' => {
                        json.push_str("\\u00");
                        for _ in 0..2 {
                            let hex = *self.chars.get(self.pos).ok_or("无效十六进制转义")?;
                            if !hex.is_ascii_hexdigit() {
                                return Err("无效十六进制转义".into());
                            }
                            json.push(hex);
                            self.pos += 1;
                        }
                    }
                    '"' | '\\' | '/' | 'b' | 'f' | 'n' | 'r' | 't' | 'u' => {
                        json.push('\\');
                        json.push(next);
                    }
                    _ => return Err("不支持的字符串转义".into()),
                }
            } else {
                if ch.is_control() {
                    return Err("字符串不能包含未转义控制字符".into());
                }
                if ch == '"' {
                    json.push('\\');
                }
                json.push(ch);
            }
        }
        Err("字符串未闭合".into())
    }
    fn key(&mut self) -> Result<String, String> {
        self.whitespace();
        if matches!(self.chars.get(self.pos), Some('\'' | '"')) {
            return self.string();
        }
        let start = self.pos;
        while self
            .chars
            .get(self.pos)
            .is_some_and(|c| c.is_ascii_alphanumeric() || *c == '_' || *c == '$')
        {
            self.pos += 1;
        }
        let key: String = self.chars[start..self.pos].iter().collect();
        if key.is_empty() || key.starts_with(|c: char| c.is_ascii_digit()) {
            return Err("无效对象键".into());
        }
        Ok(key)
    }
    fn value(&mut self, depth: usize) -> Result<Value, String> {
        self.nodes += 1;
        if depth > 16 || self.nodes > 4096 {
            return Err("fetch 数据嵌套或项目数量超过限制".into());
        }
        self.whitespace();
        match self.chars.get(self.pos).copied() {
            Some('\'' | '"') => self.string().map(Value::String),
            Some('{') => {
                self.pos += 1;
                let mut map = Map::new();
                if self.consume('}') {
                    return Ok(Value::Object(map));
                }
                loop {
                    let key = self.key()?;
                    self.punctuation(':')?;
                    let value = self.value(depth + 1)?;
                    if map.insert(key, value).is_some() {
                        return Err("fetch 对象包含重复键".into());
                    }
                    if self.consume('}') {
                        break;
                    }
                    self.punctuation(',')?;
                    if self.consume('}') {
                        break;
                    }
                }
                Ok(Value::Object(map))
            }
            Some('[') => {
                self.pos += 1;
                let mut values = Vec::new();
                if self.consume(']') {
                    return Ok(Value::Array(values));
                }
                loop {
                    values.push(self.value(depth + 1)?);
                    if self.consume(']') {
                        break;
                    }
                    self.punctuation(',')?;
                    if self.consume(']') {
                        break;
                    }
                }
                Ok(Value::Array(values))
            }
            Some('t') => {
                self.literal("true")?;
                Ok(Value::Bool(true))
            }
            Some('f') => {
                self.literal("false")?;
                Ok(Value::Bool(false))
            }
            Some('n') => {
                self.literal("null")?;
                Ok(Value::Null)
            }
            Some('-' | '0'..='9') => {
                let start = self.pos;
                while self
                    .chars
                    .get(self.pos)
                    .is_some_and(|c| c.is_ascii_digit() || ".eE+-".contains(*c))
                {
                    self.pos += 1;
                }
                let raw: String = self.chars[start..self.pos].iter().collect();
                serde_json::from_str(&raw).map_err(|_| "无效数字".into())
            }
            _ => Err("仅支持数据字面量，不支持 JavaScript 表达式、函数或变量".into()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn realistic_chrome_fetch_and_single_quotes() {
        let request = parse(r#"fetch("https://example.com/api?q=a,b", {
            "headers": {"accept": "application/json", "content-type": "application/json", "sec-fetch-mode": "cors"},
            "referrer": "https://example.com/", "body": "{\"id\":9007199254740993}",
            "method": "POST", "mode": "cors", "credentials": "include"
        });"#).unwrap();
        assert_eq!(request.method, Method::POST);
        assert_eq!(request.body.as_deref(), Some(r#"{"id":9007199254740993}"#));
        let single = parse(r#"fetch ('https://example.com/a\x2fb', {headers: {'x-note': 'a\'b "c"'}, method: 'POST', body: '',});"#).unwrap();
        assert_eq!(single.url.path(), "/a/b");
        assert_eq!(single.headers["x-note"], "a'b \"c\"");
        assert_eq!(single.body.as_deref(), Some(""));
        assert_eq!(
            parse("fetch('https://example.com')").unwrap().method,
            Method::GET
        );
    }
    #[test]
    fn curl_literal_data_and_restricted_arguments() {
        let request = parse(
            "curl 'https://example.com' \\\n -H 'X-Test: yes' --data-raw '{\"a\":1}' --compressed",
        )
        .unwrap();
        assert_eq!(request.method, Method::POST);
        assert_eq!(request.headers["x-test"], "yes");
        assert_eq!(request.body.as_deref(), Some("{\"a\":1}"));
        assert_eq!(
            parse("curl https://example.com --data-raw ''")
                .unwrap()
                .body
                .as_deref(),
            Some("")
        );
        for command in [
            "curl file:///tmp/data",
            "curl https://example.com --output a",
            "curl https://example.com --config a",
            "curl https://example.com --data @a",
            "curl https://example.com --next https://example.org",
            "curl https://example.com --location",
            "curl https://example.com 'extra'",
            "curl https://example.com -H 'Host: other'",
            "curl https://user:pass@example.com",
            "curl 'https://example.com",
            "curl https://example.com \\",
        ] {
            assert!(parse(command).is_err(), "{command}");
        }
    }
    #[test]
    fn malformed_or_non_data_fetch_is_rejected() {
        for command in [
            "fetch('https://example.com', options)",
            "fetch('https://example.com', {body: undefined})",
            "fetch('https://example.com', {method:'POST', body: JSON.stringify({a:1})})",
            "fetch('https://example.com', {method:'GET', body:'x'})",
            "fetch('https://example.com', {headers:{a:'b',a:'c'}})",
            "fetch('https://example.com', {}) trailing",
            "fetch('https://example.com'",
            "fetch(`https://example.com`)",
            "fetch('https://example.com', {unknown: 'x'})",
        ] {
            assert!(parse(command).is_err(), "{command}");
        }
        assert!(parse(&"x".repeat(MAX_COMMAND + 1)).is_err());
        let nested = format!(
            "fetch('https://example.com', {{headers: {}0{}}})",
            "[".repeat(18),
            "]".repeat(18)
        );
        assert!(parse(&nested).is_err());
    }

    fn server(response: Vec<u8>) -> (String, std::thread::JoinHandle<()>) {
        use std::io::{Read, Write};
        let listener = std::net::TcpListener::bind("127.0.0.1:0").unwrap();
        let url = format!("http://{}/data", listener.local_addr().unwrap());
        let thread = std::thread::spawn(move || {
            let (mut stream, _) = listener.accept().unwrap();
            stream
                .set_read_timeout(Some(Duration::from_secs(3)))
                .unwrap();
            let mut request = [0; 4096];
            let _ = stream.read(&mut request);
            let _ = stream.write_all(&response);
        });
        (url, thread)
    }

    #[test]
    fn bounded_real_http_and_error_paths() {
        let json = r#"[{"id":9007199254740993}]"#;
        let (url, thread) = server(
            format!(
                "HTTP/1.1 200 OK\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{json}",
                json.len()
            )
            .into_bytes(),
        );
        assert_eq!(
            tauri::async_runtime::block_on(execute(&format!("fetch('{url}')"))).unwrap(),
            json
        );
        thread.join().unwrap();
        for response in [
            b"HTTP/1.1 302 Found\r\nLocation: http://127.0.0.1:1\r\nContent-Length: 0\r\n\r\n"
                .to_vec(),
            b"HTTP/1.1 500 Error\r\nContent-Length: 0\r\n\r\n".to_vec(),
            b"HTTP/1.1 200 OK\r\nContent-Length: 3\r\n\r\nno!".to_vec(),
            b"HTTP/1.1 200 OK\r\nContent-Length: 1\r\n\r\n\xff".to_vec(),
            format!(
                "HTTP/1.1 200 OK\r\nContent-Length: {}\r\n\r\n",
                MAX_RESPONSE + 1
            )
            .into_bytes(),
        ] {
            let (url, thread) = server(response);
            assert!(tauri::async_runtime::block_on(execute(&format!("curl '{url}'"))).is_err());
            thread.join().unwrap();
        }
        // No Content-Length: enforce the bound while streaming as well.
        let mut response = b"HTTP/1.0 200 OK\r\n\r\n".to_vec();
        response.resize(response.len() + MAX_RESPONSE + 1, b' ');
        let (url, thread) = server(response);
        let error = tauri::async_runtime::block_on(execute(&format!("curl '{url}'"))).unwrap_err();
        assert!(error.contains("16 MiB"));
        thread.join().unwrap();

        use std::io::{Read, Write};
        let listener = std::net::TcpListener::bind("127.0.0.1:0").unwrap();
        let url = format!("http://{}", listener.local_addr().unwrap());
        let thread = std::thread::spawn(move || {
            let (mut stream, _) = listener.accept().unwrap();
            let _ = stream.read(&mut [0; 1024]);
            std::thread::sleep(Duration::from_millis(200));
            let _ = stream.write_all(b"HTTP/1.1 200 OK\r\nContent-Length: 2\r\n\r\n[]");
        });
        assert!(
            tauri::async_runtime::block_on(send(parse(&url).unwrap(), Duration::from_millis(50)))
                .is_err()
        );
        thread.join().unwrap();

        // Busy guard releases after every error, without queuing unlimited requests.
        REQUEST_ACTIVE.store(true, Ordering::Release);
        assert!(
            tauri::async_runtime::block_on(execute("invalid"))
                .unwrap_err()
                .contains("已有请求")
        );
        REQUEST_ACTIVE.store(false, Ordering::Release);
        assert!(tauri::async_runtime::block_on(execute("invalid")).is_err());
        assert!(!REQUEST_ACTIVE.load(Ordering::Acquire));
    }
}
