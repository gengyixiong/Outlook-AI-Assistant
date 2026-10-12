# Local Semantic Search Service

This directory contains the local-only embedding and retrieval service used by
the Outlook AI Assistant add-in. It runs `google/embeddinggemma-2` on this
machine and owns the SQLite index of your mail. No mail content leaves the
computer.

## Quick start

### Recommended: install to a stable location first

The release ZIP is often extracted into Downloads, which Windows cleanup may
delete. Installing the service to a fixed per-user location means the
extracted folder can then be removed safely:

```text
install-service.cmd   # copies the service to %%LOCALAPPDATA%% and makes a shortcut
```

It installs to:

```text
%LOCALAPPDATA%\OutlookAiAssistant\embedding_service
```

Uninstall it with `uninstall-service.ps1` (from the release folder). By
default it keeps the downloaded model and the semantic index so a reinstall
does not re-download ~1.5 GB or rebuild the index; pass `-RemoveModel` and/or
`-RemoveIndex` to delete them too.

### Then run these in order

The model download step is **required before indexing**; without it the
add-in reports that the model is not ready instead of building a useless
index.

```text
setup-venv.cmd        # create .venv and install dependencies (once)
download-model.cmd    # download and cache the model (once, ~1.5 GB, needs network)
start-service.cmd     # start the service on 127.0.0.1:8765
```

Run them from the installed folder (`install-service.cmd` creates a Start
Menu shortcut for `start-service.cmd`). Each script keeps its window open at
the end so you can read the result.
`download-model.cmd` shows live progress (`Fetching N files: xx%` and the
download speed) and prints `SUCCESS` or a clear failure reason. On a machine
that already has the model cached it finishes in seconds.

Then, in Outlook, open the **Semantic Search** tab and click
**Download model** if the status says the model is not ready, then
**Build / update semantic index**.

The add-in reads the shared token from
`%LOCALAPPDATA%\OutlookAiAssistant\semantic-service-token` and sends it as a
bearer token. The token is created automatically on first start.

### Prerequisites

- Python 3.10–3.12 must be installed from python.org with
  **Add python.exe to PATH** enabled.
- If `setup-venv.cmd` reports that Python was not found, or shows a Microsoft
  Store message, the `python` command is pointing at the Store stub. Disable
  the `python.exe` / `python3.exe` aliases under
  Settings > Apps > Advanced app settings > App execution aliases, then retry.
  `setup-venv.cmd` prefers the `py` launcher and skips the Store stub on
  PATH, and also probes the usual install directories.

### Why the download step is separate

`google/embeddinggemma-2` is not bundled with the release (it is several GB).
On a machine that has never downloaded it, indexing would fail for every
mail. Rather than silently writing vector-less rows, the service refuses to
index and the add-in tells you to download the model first. You can trigger
the same download from the Outlook UI (**Download model**) or from the
command line:

```text
download-model.cmd
```

## Why a separate service?

`google/embeddinggemma-2` is a PyTorch model. Keeping it in a Python process
means:

* the C# add-in stays on .NET Framework 4.8 with no NuGet dependencies;
* model loading and inference never block the Outlook UI thread;
* the service can be started, stopped, or restarted independently, and its
  absence never crashes Outlook.

## Storage layout

Code and data are deliberately separated:

| Path | Contents | Safe to delete? |
| --- | --- | --- |
| Extracted release folder | install scripts + service source | Yes, after `install-service.cmd` |
| `%LOCALAPPDATA%\OutlookAiAssistant\embedding_service` | installed service code + `.venv` | Only via `uninstall-service.ps1` |
| `%LOCALAPPDATA%\OutlookAiAssistant\semantic-index\mail.db` | SQLite metadata, chunks, vectors, FTS index | Yes (also clears the index) |
| `%LOCALAPPDATA%\OutlookAiAssistant\models\` | Hugging Face model cache | Yes (re-download needed) |
| `%LOCALAPPDATA%\OutlookAiAssistant\semantic-service-token` | loopback bearer token | Regenerated on start |
| `%LOCALAPPDATA%\OutlookAiAssistant\logs\semantic-service.log` | service log (no mail content) | Yes |

The Python service never writes anything next to its code, so the code
location and the data location can move independently.

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
