/// llm_client.rs
/// 다중 공급자 LLM 에이전틱 클라이언트 (OpenAI / Anthropic / Gemini)
/// Features: cancellation, SSE streaming (OpenAI/Anthropic), token tracking + pruning

use crate::commands::HistoryMessage;
use crate::tools::{hwp_tools, to_anthropic_tools, to_gemini_tools, to_openai_tools, SYSTEM_PROMPT};
use anyhow::Context;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::{pin::Pin, sync::Arc};
use tokio_util::sync::CancellationToken;

pub type BoxFuture<T> = Pin<Box<dyn std::future::Future<Output = T> + Send + 'static>>;
pub type ToolExecutor = Arc<dyn Fn(String, Value) -> BoxFuture<String> + Send + Sync + 'static>;
pub type EventCallback = Arc<dyn Fn(AgentEvent) + Send + Sync + 'static>;

pub const DEFAULT_MODELS: &[(&str, &str)] = &[
    ("openai", "gpt-4o"),
    ("gemini", "gemini-2.0-flash"),
    ("anthropic", "claude-sonnet-4-6"),
];

pub fn default_model(provider: &str) -> &'static str {
    DEFAULT_MODELS
        .iter()
        .find(|(p, _)| *p == provider)
        .map(|(_, m)| *m)
        .unwrap_or("gpt-4o")
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum AgentEvent {
    ToolCall { name: String, args: Value },
    ToolResult { name: String, result: String },
    ToolConfirmRequest { name: String, args: Value },
    LlmThinking { text: String },
    FinalResponse { text: String },
    Error { message: String },
    TokenUsage {
        prompt_tokens: u64,
        completion_tokens: u64,
        total_tokens: u64,
    },
}

pub struct LlmClient {
    pub provider: String,
    pub api_key: String,
    pub model: String,
    http: reqwest::Client,
}

// ──────────────────────────────────────────────────────────
// Context window helpers
// ──────────────────────────────────────────────────────────

fn context_limit(model: &str) -> u64 {
    match model {
        m if m.starts_with("gpt-4o") => 128_000,
        m if m.starts_with("gpt-4") => 128_000,
        m if m.starts_with("o1") || m.starts_with("o3") || m.starts_with("o4") => 200_000,
        m if m.contains("claude") => 200_000,
        m if m.contains("gemini") => 1_000_000,
        _ => 128_000,
    }
}

/// Truncate old tool-result messages when approaching context limit.
/// Keeps the most recent `keep_recent` tool-result entries intact.
/// 오래된 tool result를 항상 80자로 잘라 TPM 절감 (keep_recent개만 원본 보존)
fn prune_openai_messages(messages: &mut [Value], keep_recent: usize) {
    let tool_indices: Vec<usize> = messages
        .iter()
        .enumerate()
        .filter(|(_, m)| m["role"] == "tool")
        .map(|(i, _)| i)
        .collect();
    if tool_indices.len() <= keep_recent {
        return;
    }
    let cutoff = tool_indices.len() - keep_recent;
    for &idx in &tool_indices[..cutoff] {
        if let Some(content) = messages[idx]["content"].as_str() {
            if content.len() > 80 {
                let safe = content.char_indices().nth(80).map(|(i, _)| i).unwrap_or(content.len());
                messages[idx]["content"] = json!(format!("{}…[생략]", &content[..safe]));
            }
        }
    }
}

fn prune_anthropic_messages(messages: &mut [Value], keep_recent: usize) {
    let mut tool_result_locs: Vec<(usize, usize)> = vec![];
    for (mi, msg) in messages.iter().enumerate() {
        if let Some(arr) = msg["content"].as_array() {
            for (bi, block) in arr.iter().enumerate() {
                if block["type"] == "tool_result" {
                    tool_result_locs.push((mi, bi));
                }
            }
        }
    }
    if tool_result_locs.len() <= keep_recent {
        return;
    }
    let cutoff = tool_result_locs.len() - keep_recent;
    for &(mi, bi) in &tool_result_locs[..cutoff] {
        if let Some(content) = messages[mi]["content"][bi]["content"].as_str() {
            if content.len() > 80 {
                let safe = content.char_indices().nth(80).map(|(i, _)| i).unwrap_or(content.len());
                messages[mi]["content"][bi]["content"] =
                    json!(format!("{}…[생략]", &content[..safe]));
            }
        }
    }
}

fn prune_gemini_contents(contents: &mut [Value], keep_recent: usize) {
    let mut fn_resp_locs: Vec<(usize, usize)> = vec![];
    for (ci, content) in contents.iter().enumerate() {
        if let Some(parts) = content["parts"].as_array() {
            for (pi, part) in parts.iter().enumerate() {
                if part.get("functionResponse").is_some() {
                    fn_resp_locs.push((ci, pi));
                }
            }
        }
    }
    if fn_resp_locs.len() <= keep_recent {
        return;
    }
    let cutoff = fn_resp_locs.len() - keep_recent;
    for &(ci, pi) in &fn_resp_locs[..cutoff] {
        if let Some(result) = contents[ci]["parts"][pi]["functionResponse"]["response"]["result"]
            .as_str()
        {
            if result.len() > 80 {
                let safe = result.char_indices().nth(80).map(|(i, _)| i).unwrap_or(result.len());
                contents[ci]["parts"][pi]["functionResponse"]["response"]["result"] =
                    json!(format!("{}…[생략]", &result[..safe]));
            }
        }
    }
}

// ──────────────────────────────────────────────────────────
// SSE parsing helpers
// ──────────────────────────────────────────────────────────

/// Parse SSE lines from a byte chunk, returning (event_type, data) pairs.
/// Incomplete trailing line is kept in `buffer` for the next call.
fn parse_sse_lines(buffer: &mut String, chunk: &[u8]) -> Vec<(Option<String>, String)> {
    buffer.push_str(&String::from_utf8_lossy(chunk));
    let mut events = vec![];
    let mut current_event: Option<String> = None;

    // Process complete lines (split by \n)
    while let Some(newline_pos) = buffer.find('\n') {
        let line = buffer[..newline_pos].trim_end_matches('\r').to_string();
        buffer.drain(..=newline_pos);

        if line.is_empty() {
            // Empty line = event boundary (but we handle per data: line)
            continue;
        }
        if let Some(ev) = line.strip_prefix("event: ") {
            current_event = Some(ev.to_string());
        } else if let Some(data) = line.strip_prefix("data: ") {
            events.push((current_event.take(), data.to_string()));
        }
        // Ignore other lines (id:, retry:, comments)
    }
    events
}

// ──────────────────────────────────────────────────────────
// LlmClient
// ──────────────────────────────────────────────────────────

impl LlmClient {
    pub fn new(provider: &str, api_key: &str, model: Option<&str>) -> anyhow::Result<Self> {
        let cleaned: String = api_key
            .chars()
            .filter(|c| c.is_ascii() && !c.is_control())
            .collect();
        if cleaned.is_empty() {
            anyhow::bail!("API Key가 비어 있거나 유효하지 않습니다.");
        }
        let model = model
            .filter(|m| !m.is_empty())
            .map(str::to_string)
            .unwrap_or_else(|| default_model(provider).to_string());

        let http = reqwest::Client::builder()
            .timeout(std::time::Duration::from_secs(120))
            .build()?;

        Ok(LlmClient {
            provider: provider.to_string(),
            api_key: cleaned,
            model,
            http,
        })
    }

    pub async fn call_agentic(
        &self,
        doc_summary: &str,
        user_query: &str,
        history: &[HistoryMessage],
        tool_executor: ToolExecutor,
        event_cb: EventCallback,
        max_turns: usize,
        cancel_token: CancellationToken,
    ) -> anyhow::Result<String> {
        let initial = format!(
            "[문서 초기 정보]\n{doc_summary}\n\n[사용자 요청]\n{user_query}"
        );
        match self.provider.as_str() {
            "openai" => {
                self.agentic_openai(&initial, history, tool_executor, event_cb, max_turns, cancel_token)
                    .await
            }
            "anthropic" => {
                self.agentic_anthropic(&initial, history, tool_executor, event_cb, max_turns, cancel_token)
                    .await
            }
            "gemini" => {
                self.agentic_gemini(&initial, history, tool_executor, event_cb, max_turns, cancel_token)
                    .await
            }
            p => anyhow::bail!("알 수 없는 공급자: {p}"),
        }
    }

    // ─────────────── OpenAI (streaming) ────────────────────

    async fn agentic_openai(
        &self,
        initial_message: &str,
        history: &[HistoryMessage],
        tool_executor: ToolExecutor,
        event_cb: EventCallback,
        max_turns: usize,
        cancel_token: CancellationToken,
    ) -> anyhow::Result<String> {
        let tools = to_openai_tools(&hwp_tools());
        let mut messages: Vec<Value> =
            vec![json!({"role": "system", "content": SYSTEM_PROMPT})];
        for h in history {
            messages.push(json!({"role": h.role, "content": h.content}));
        }
        messages.push(json!({"role": "user", "content": initial_message}));

        let limit = context_limit(&self.model);
        let mut cumulative_tokens: u64 = 0;

        for _turn in 0..max_turns {
            if cancel_token.is_cancelled() {
                anyhow::bail!("에이전트가 취소되었습니다.");
            }

            // 매 턴마다 오래된 tool result 축약 (TPM 절감); 최근 3개만 원본 유지
            prune_openai_messages(&mut messages, 3);
            // 컨텍스트 한계 근접 시 더 공격적으로 정리
            if cumulative_tokens > limit * 80 / 100 {
                prune_openai_messages(&mut messages, 1);
            }

            let body = json!({
                "model": self.model,
                "messages": messages,
                "tools": tools,
                "tool_choice": "auto",
                "temperature": 0.1,
                "stream": true,
                "stream_options": {"include_usage": true},
            });

            let resp = tokio::select! {
                r = self.http
                    .post("https://api.openai.com/v1/chat/completions")
                    .header("Authorization", format!("Bearer {}", self.api_key))
                    .json(&body)
                    .send() => {
                    r.context("OpenAI API 요청 실패")?
                }
                _ = cancel_token.cancelled() => {
                    anyhow::bail!("에이전트가 취소되었습니다.");
                }
            };

            if !resp.status().is_success() {
                let status = resp.status();
                let text = resp.text().await.unwrap_or_default();
                anyhow::bail!("OpenAI API 오류 {status}: {text}");
            }

            // Stream SSE chunks
            let mut sse_buf = String::new();
            let mut text_acc = String::new();
            let mut finish_reason = String::new();
            // tool_calls accumulator: index -> (id, name, arguments_str)
            let mut tc_acc: Vec<(String, String, String)> = vec![];
            let mut usage_data: Option<(u64, u64, u64)> = None;

            let mut stream = resp;
            loop {
                let chunk_result = tokio::select! {
                    c = stream.chunk() => c,
                    _ = cancel_token.cancelled() => {
                        anyhow::bail!("에이전트가 취소되었습니다.");
                    }
                };
                let chunk = match chunk_result.context("OpenAI 스트림 읽기 실패")? {
                    Some(c) => c,
                    None => break, // stream ended
                };

                let events = parse_sse_lines(&mut sse_buf, &chunk);
                for (_evt, data) in events {
                    if data == "[DONE]" {
                        break;
                    }
                    let parsed: Value = match serde_json::from_str(&data) {
                        Ok(v) => v,
                        Err(_) => continue,
                    };

                    // Extract usage from final chunk
                    if let Some(usage) = parsed.get("usage") {
                        let pt = usage["prompt_tokens"].as_u64().unwrap_or(0);
                        let ct = usage["completion_tokens"].as_u64().unwrap_or(0);
                        let tt = usage["total_tokens"].as_u64().unwrap_or(0);
                        usage_data = Some((pt, ct, tt));
                    }

                    let delta = &parsed["choices"][0]["delta"];
                    if delta.is_null() {
                        continue;
                    }

                    // finish_reason
                    if let Some(fr) = parsed["choices"][0]["finish_reason"].as_str() {
                        finish_reason = fr.to_string();
                    }

                    // Text content delta
                    if let Some(content) = delta["content"].as_str() {
                        if !content.is_empty() {
                            text_acc.push_str(content);
                            event_cb(AgentEvent::LlmThinking {
                                text: content.to_string(),
                            });
                        }
                    }

                    // Tool calls delta
                    if let Some(tcs) = delta["tool_calls"].as_array() {
                        for tc in tcs {
                            let idx = tc["index"].as_u64().unwrap_or(0) as usize;
                            while tc_acc.len() <= idx {
                                tc_acc.push((String::new(), String::new(), String::new()));
                            }
                            if let Some(id) = tc["id"].as_str() {
                                tc_acc[idx].0 = id.to_string();
                            }
                            if let Some(name) = tc["function"]["name"].as_str() {
                                tc_acc[idx].1.push_str(name);
                            }
                            if let Some(args) = tc["function"]["arguments"].as_str() {
                                tc_acc[idx].2.push_str(args);
                            }
                        }
                    }
                }
            }

            // Emit token usage
            if let Some((pt, ct, tt)) = usage_data {
                cumulative_tokens = tt;
                event_cb(AgentEvent::TokenUsage {
                    prompt_tokens: pt,
                    completion_tokens: ct,
                    total_tokens: tt,
                });
            }

            // Build assistant message for history
            if !tc_acc.is_empty() {
                let tool_calls_json: Vec<Value> = tc_acc
                    .iter()
                    .map(|(id, name, args)| {
                        json!({
                            "id": id,
                            "type": "function",
                            "function": {"name": name, "arguments": args}
                        })
                    })
                    .collect();
                let mut msg = json!({"role": "assistant"});
                if !text_acc.is_empty() {
                    msg["content"] = json!(text_acc);
                } else {
                    msg["content"] = Value::Null;
                }
                msg["tool_calls"] = json!(tool_calls_json);
                messages.push(msg);

                // Execute tool calls
                let mut results = vec![];
                for (id, name, args_str) in &tc_acc {
                    if cancel_token.is_cancelled() {
                        anyhow::bail!("에이전트가 취소되었습니다.");
                    }
                    let args: Value =
                        serde_json::from_str(args_str).unwrap_or(json!({}));

                    // ToolCall 이벤트는 tool_executor 안에서 확인 후 발생 (write tool 타이밍 정합성)
                    let result = tool_executor(name.clone(), args).await;
                    event_cb(AgentEvent::ToolResult {
                        name: name.clone(),
                        result: result.clone(),
                    });

                    results.push(json!({
                        "role": "tool",
                        "tool_call_id": id,
                        "content": result,
                    }));
                }
                messages.extend(results);
                continue;
            }

            // No tool calls — text-only response
            if finish_reason == "stop" || !text_acc.is_empty() {
                return Ok(text_acc);
            }
        }
        Ok("⚠️ 최대 턴 수 초과".to_string())
    }

    // ─────────────── Anthropic (streaming) ──────────────────

    async fn agentic_anthropic(
        &self,
        initial_message: &str,
        history: &[HistoryMessage],
        tool_executor: ToolExecutor,
        event_cb: EventCallback,
        max_turns: usize,
        cancel_token: CancellationToken,
    ) -> anyhow::Result<String> {
        let tools = to_anthropic_tools(&hwp_tools());
        let mut messages: Vec<Value> = vec![];
        for h in history {
            messages.push(json!({"role": h.role, "content": h.content}));
        }
        messages.push(json!({"role": "user", "content": initial_message}));

        let limit = context_limit(&self.model);
        let mut cumulative_tokens: u64 = 0;

        for _turn in 0..max_turns {
            if cancel_token.is_cancelled() {
                anyhow::bail!("에이전트가 취소되었습니다.");
            }

            prune_anthropic_messages(&mut messages, 3);
            if cumulative_tokens > limit * 80 / 100 {
                prune_anthropic_messages(&mut messages, 1);
            }

            let body = json!({
                "model": self.model,
                "max_tokens": 8192,
                "system": SYSTEM_PROMPT,
                "tools": tools,
                "messages": messages,
                "stream": true,
            });

            let resp = tokio::select! {
                r = self.http
                    .post("https://api.anthropic.com/v1/messages")
                    .header("x-api-key", &self.api_key)
                    .header("anthropic-version", "2023-06-01")
                    .json(&body)
                    .send() => {
                    r.context("Anthropic API 요청 실패")?
                }
                _ = cancel_token.cancelled() => {
                    anyhow::bail!("에이전트가 취소되었습니다.");
                }
            };

            if !resp.status().is_success() {
                let status = resp.status();
                let text = resp.text().await.unwrap_or_default();
                anyhow::bail!("Anthropic API 오류 {status}: {text}");
            }

            // Stream SSE
            let mut sse_buf = String::new();
            let mut stop_reason = String::new();
            // Content blocks: (type, text_or_json, tool_id, tool_name)
            let mut blocks: Vec<(String, String, String, String)> = vec![];
            let mut current_block_idx: Option<usize> = None;
            let mut usage_data: Option<(u64, u64, u64)> = None;

            let mut stream = resp;
            loop {
                let chunk_result = tokio::select! {
                    c = stream.chunk() => c,
                    _ = cancel_token.cancelled() => {
                        anyhow::bail!("에이전트가 취소되었습니다.");
                    }
                };
                let chunk = match chunk_result.context("Anthropic 스트림 읽기 실패")? {
                    Some(c) => c,
                    None => break,
                };

                let events = parse_sse_lines(&mut sse_buf, &chunk);
                for (evt_type, data) in events {
                    let evt = evt_type.as_deref().unwrap_or("");
                    let parsed: Value = match serde_json::from_str(&data) {
                        Ok(v) => v,
                        Err(_) => continue,
                    };

                    match evt {
                        "message_start" => {
                            // Extract input token count from message_start
                            if let Some(usage) = parsed["message"].get("usage") {
                                let input = usage["input_tokens"].as_u64().unwrap_or(0);
                                usage_data = Some((input, 0, input));
                            }
                        }
                        "content_block_start" => {
                            let cb = &parsed["content_block"];
                            let btype = cb["type"].as_str().unwrap_or("").to_string();
                            let tool_id =
                                cb["id"].as_str().unwrap_or("").to_string();
                            let tool_name =
                                cb["name"].as_str().unwrap_or("").to_string();
                            blocks.push((btype, String::new(), tool_id, tool_name));
                            current_block_idx = Some(blocks.len() - 1);
                        }
                        "content_block_delta" => {
                            if let Some(idx) = current_block_idx {
                                let delta = &parsed["delta"];
                                let delta_type =
                                    delta["type"].as_str().unwrap_or("");
                                match delta_type {
                                    "text_delta" => {
                                        let text =
                                            delta["text"].as_str().unwrap_or("");
                                        if !text.is_empty() {
                                            blocks[idx].1.push_str(text);
                                            event_cb(AgentEvent::LlmThinking {
                                                text: text.to_string(),
                                            });
                                        }
                                    }
                                    "input_json_delta" => {
                                        let json_frag = delta["partial_json"]
                                            .as_str()
                                            .unwrap_or("");
                                        blocks[idx].1.push_str(json_frag);
                                    }
                                    _ => {}
                                }
                            }
                        }
                        "content_block_stop" => {
                            current_block_idx = None;
                        }
                        "message_delta" => {
                            if let Some(sr) = parsed["delta"]["stop_reason"].as_str()
                            {
                                stop_reason = sr.to_string();
                            }
                            if let Some(usage) = parsed.get("usage") {
                                let output =
                                    usage["output_tokens"].as_u64().unwrap_or(0);
                                if let Some(ref mut ud) = usage_data {
                                    ud.1 = output;
                                    ud.2 = ud.0 + output;
                                }
                            }
                        }
                        "message_stop" => {
                            // Stream complete
                        }
                        _ => {}
                    }
                }
            }

            // Emit token usage
            if let Some((pt, ct, tt)) = usage_data {
                cumulative_tokens = tt;
                event_cb(AgentEvent::TokenUsage {
                    prompt_tokens: pt,
                    completion_tokens: ct,
                    total_tokens: tt,
                });
            }

            // Reconstruct content array for message history
            let content_blocks: Vec<Value> = blocks
                .iter()
                .map(|(btype, text, id, name)| {
                    if btype == "tool_use" {
                        let input: Value =
                            serde_json::from_str(text).unwrap_or(json!({}));
                        json!({"type": "tool_use", "id": id, "name": name, "input": input})
                    } else {
                        json!({"type": "text", "text": text})
                    }
                })
                .collect();

            messages.push(json!({"role": "assistant", "content": content_blocks}));

            if stop_reason == "end_turn" {
                // Return accumulated text
                let text = blocks
                    .iter()
                    .find(|(t, _, _, _)| t == "text")
                    .map(|(_, s, _, _)| s.clone())
                    .unwrap_or_default();
                return Ok(text);
            }

            if stop_reason == "tool_use" {
                let mut tool_results = vec![];
                for (btype, json_str, id, name) in &blocks {
                    if btype != "tool_use" {
                        continue;
                    }
                    if cancel_token.is_cancelled() {
                        anyhow::bail!("에이전트가 취소되었습니다.");
                    }
                    let args: Value =
                        serde_json::from_str(json_str).unwrap_or(json!({}));

                    // ToolCall 이벤트는 tool_executor 안에서 확인 후 발생
                    let result = tool_executor(name.clone(), args).await;
                    event_cb(AgentEvent::ToolResult {
                        name: name.clone(),
                        result: result.clone(),
                    });

                    tool_results.push(json!({
                        "type": "tool_result",
                        "tool_use_id": id,
                        "content": result,
                    }));
                }
                messages.push(json!({"role": "user", "content": tool_results}));
                continue;
            }

            // Fallback: return any text
            let text = blocks
                .iter()
                .find(|(t, _, _, _)| t == "text")
                .map(|(_, s, _, _)| s.clone())
                .unwrap_or_default();
            if !text.is_empty() {
                return Ok(text);
            }
        }
        Ok("⚠️ 최대 턴 수 초과".to_string())
    }

    // ─────────────── Gemini (non-streaming) ─────────────────

    async fn agentic_gemini(
        &self,
        initial_message: &str,
        history: &[HistoryMessage],
        tool_executor: ToolExecutor,
        event_cb: EventCallback,
        max_turns: usize,
        cancel_token: CancellationToken,
    ) -> anyhow::Result<String> {
        let tools = to_gemini_tools(&hwp_tools());
        let url = format!(
            "https://generativelanguage.googleapis.com/v1beta/models/{}:generateContent?key={}",
            self.model, self.api_key
        );
        let mut contents: Vec<Value> = vec![];
        for h in history {
            let gemini_role = if h.role == "assistant" { "model" } else { "user" };
            contents.push(json!({"role": gemini_role, "parts": [{"text": h.content}]}));
        }
        contents.push(json!({"role": "user", "parts": [{"text": initial_message}]}));
        let system_instruction = json!({"parts": [{"text": SYSTEM_PROMPT}]});

        let limit = context_limit(&self.model);
        let mut cumulative_tokens: u64 = 0;

        for _turn in 0..max_turns {
            if cancel_token.is_cancelled() {
                anyhow::bail!("에이전트가 취소되었습니다.");
            }

            prune_gemini_contents(&mut contents, 3);
            if cumulative_tokens > limit * 80 / 100 {
                prune_gemini_contents(&mut contents, 1);
            }

            let body = json!({
                "system_instruction": system_instruction,
                "tools": tools,
                "generation_config": {"temperature": 0.1},
                "contents": contents,
            });

            let resp = tokio::select! {
                r = self.http.post(&url).json(&body).send() => {
                    r.context("Gemini API 요청 실패")?
                }
                _ = cancel_token.cancelled() => {
                    anyhow::bail!("에이전트가 취소되었습니다.");
                }
            };

            if !resp.status().is_success() {
                let status = resp.status();
                let text = resp.text().await.unwrap_or_default();
                anyhow::bail!("Gemini API 오류 {status}: {text}");
            }

            let data: Value = resp.json().await.context("Gemini 응답 파싱 실패")?;

            // Extract token usage
            if let Some(meta) = data.get("usageMetadata") {
                let pt = meta["promptTokenCount"].as_u64().unwrap_or(0);
                let ct = meta["candidatesTokenCount"].as_u64().unwrap_or(0);
                let tt = meta["totalTokenCount"].as_u64().unwrap_or(0);
                cumulative_tokens = tt;
                event_cb(AgentEvent::TokenUsage {
                    prompt_tokens: pt,
                    completion_tokens: ct,
                    total_tokens: tt,
                });
            }

            let parts = data["candidates"][0]["content"]["parts"]
                .as_array()
                .cloned()
                .unwrap_or_default();

            // Emit any text as thinking
            for p in &parts {
                if let Some(text) = p["text"].as_str() {
                    if !text.is_empty() {
                        event_cb(AgentEvent::LlmThinking {
                            text: text.to_string(),
                        });
                    }
                }
            }

            let fn_calls: Vec<Value> = parts
                .iter()
                .filter(|p| p.get("functionCall").is_some())
                .cloned()
                .collect();

            if fn_calls.is_empty() {
                let text = parts
                    .iter()
                    .find_map(|p| p["text"].as_str())
                    .unwrap_or("")
                    .to_string();
                return Ok(text);
            }

            contents.push(json!({"role": "model", "parts": parts}));

            let mut fn_responses = vec![];
            for part in &fn_calls {
                if cancel_token.is_cancelled() {
                    anyhow::bail!("에이전트가 취소되었습니다.");
                }
                let fc = &part["functionCall"];
                let name = fc["name"].as_str().unwrap_or("").to_string();
                let args = fc["args"].clone();

                // ToolCall 이벤트는 tool_executor 안에서 확인 후 발생
                let result = tool_executor(name.clone(), args).await;
                event_cb(AgentEvent::ToolResult {
                    name: name.clone(),
                    result: result.clone(),
                });

                fn_responses.push(json!({
                    "functionResponse": {
                        "name": name,
                        "response": {"result": result}
                    }
                }));
            }
            contents.push(json!({"role": "user", "parts": fn_responses}));
        }
        Ok("⚠️ 최대 턴 수 초과".to_string())
    }
}
