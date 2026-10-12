# Local Semantic Search Service

This directory contains the local-only embedding and retrieval service used by
the Outlook AI Assistant add-in. It runs `google/embeddinggemma-2` on this
machine and owns the SQLite index of your mail. No mail content leaves the
computer.

## Quick start

```text
setup-venv.cmd        # create .venv and install dependencies (once)
download-model.cmd    # download and cache the model (once, ~1.5 GB)
start-service.cmd     # start the service on 127.0.0.1:8765
```

The add-in reads the shared token from
`%LOCALAPPDATA%\OutlookAiAssistant\semantic-service-token` and sends it as a
bearer token. The token is created automatically on first start.

## Why a separate service?

`google/embeddinggemma-2` is a PyTorch model. Keeping it in a Python process
means:

* the C# add-in stays on .NET Framework 4.8 with no NuGet dependencies;
* model loading and inference never block the Outlook UI thread;
* the service can be started, stopped, or restarted independently, and its
  absence never crashes Outlook.

## Storage layout

All per-user data lives under `%LOCALAPPDATA%\OutlookAiAssistant`:

| Path | Contents |
| --- | --- |
| `semantic-index\mail.db` | SQLite metadata, chunks, vectors, FTS index |
| `models\` | Hugging Face model cache (offline after first download) |
| `semantic-service-token` | loopback bearer token |
| `logs\semantic-service.log` | service log (no mail content) |

## Security

* Binds to `127.0.0.1` only. Never exposed to the LAN.
* Requires a bearer token, so unrelated local processes cannot read the
  database through the API.
* The token file is created with user-only permissions where supported.

## Notes on the model

* Text-only: vision and audio encoders are disabled
  (`config_kwargs={"vision_config": None, "audio_config": None}`), 270M of
  740M parameters.
* 768-dimensional, L2-normalized embeddings.
* `float32` on CPU (the safe default); `bfloat16` on supported GPUs.
  `float16` is **never** used: the model's activations exceed float16's range
  and silently produce NaN or degraded vectors.
* Query and document prompts follow the model card's asymmetric retrieval
  format, with the real email subject embedded as the document title.
* `torchvision` and `pillow` are required even for text-only use because the
  multimodal processor imports them at load time.
