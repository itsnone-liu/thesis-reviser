"""Shared LLM utilities."""
from openai import OpenAI

DEEPSEEK_BASE_URL = "https://api.deepseek.com/v1"
DEEPSEEK_API_KEY = "sk-25ef0a58ce174aa8b1e5607a433bf2c2"
_client = None

def _get_client():
    global _client
    if _client is None:
        _client = OpenAI(base_url=DEEPSEEK_BASE_URL, api_key=DEEPSEEK_API_KEY, timeout=15.0, max_retries=0)
    return _client

def deepseek_chat(messages, model="deepseek-chat", temperature=0.7, max_tokens=2000):
    client = _get_client()
    resp = client.chat.completions.create(model=model, messages=messages, temperature=temperature, max_tokens=max_tokens)
    return resp.choices[0].message.content or ""
