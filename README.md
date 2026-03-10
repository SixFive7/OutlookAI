# OutlookAI

> **Based on [OutlookAI by kirklandsig](https://github.com/kirklandsig/OutlookAI)** — originally created and released under the MIT License.

An AI-powered email writing assistant for Microsoft Outlook, built as a VSTO add-in. Uses Claude Code CLI as its AI backend, allowing you to use your existing Claude Pro or Max subscription with no separate API key or per-token billing.

> **WARNING: The screenshot below is outdated and does not reflect the current UI. It is a placeholder only and needs to be replaced with an up-to-date screenshot.**

<img width="283" height="317" alt="OutlookAI screenshot" src="https://github.com/user-attachments/assets/7513e75c-c226-4791-853a-d1aacd897883" />

---

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Claude Code Integration](#claude-code-integration)
- [Getting Started](#getting-started)
  - [Installation](#installation)
    - [Option 1: Pre-configured Build (Enterprise/RDS)](#option-1-pre-configured-build-enterpriserds)
    - [Option 2: Per-User Install](#option-2-per-user-install)
  - [Building from Source](#building-from-source)
- [Usage](#usage)
  - [Quick Actions](#quick-actions)
  - [Draft New Email](#draft-new-email)
  - [Custom Action](#custom-action)
- [Configuration](#configuration)
- [Deployment Scripts](#deployment-scripts)
- [Troubleshooting](#troubleshooting)
- [License](#license)
- [Acknowledgments](#acknowledgments)

---

## Features

- **Quick Actions** — One-click buttons to improve your email drafts:
  - Proofread (grammar, spelling, punctuation)
  - Revise (clarity and flow)
  - Shorten / Lengthen
  - Formal / Friendly tone
- **Draft New Emails** — Describe what you want to write and let AI generate the email
  - Context-aware replies (AI sees the email chain)
- **Custom Action** — Type any instruction to apply to the current email (e.g., "translate to Spanish", "add bullet points")
- **Insert, Replace, or Discard** — Review AI results and choose to insert at the top (preserving email chain), replace everything, or discard

## Requirements

| Requirement | Details |
|---|---|
| **OS** | Windows 10/11 or Windows Server 2019/2022/2025 |
| **Outlook** | Microsoft Outlook 2016, 2019, 2021, or 2024 (desktop version) |
| **Runtime** | .NET Framework 4.8 |
| **VSTO** | [Visual Studio Tools for Office Runtime](https://aka.ms/VSTORuntime) |
| **Claude Code CLI** | [Install instructions](https://code.claude.com/docs/en/getting-started) — requires a Claude Pro or Max subscription |

## Claude Code Integration

OutlookAI invokes the Claude Code CLI (`claude -p`) as a subprocess for each AI request. Several integration approaches were evaluated during development:

| Approach | Tradeoff |
|---|---|
| Persistent subprocess with NDJSON protocol | Eliminates startup latency but requires ~150 lines of protocol handling, process lifecycle management, and crash recovery |
| Claude Agent SDK via HTTP bridge | Same latency benefit, but adds a Node.js middleman process, ~150MB of deployment overhead, and an extra IPC hop |
| MCP server mode (`claude mcp serve`) | Designed for tool integration, not prompt-response — adds ~300 lines of JSON-RPC client code for no benefit in this use case |
| Fire-and-forget subprocess (`claude -p`) | Simple, but pays ~500ms CLI startup cost on every request |

OutlookAI uses a **pre-warmed fire-and-forget** approach: it spawns a `claude -p` process at Outlook startup that waits for input. When the user triggers an action, the prompt is written to the already-warm process's stdin and the response is read from stdout. A new process is immediately pre-warmed in the background for the next request. This gives the zero-latency benefit of a persistent process with the simplicity of fire-and-forget — no protocol implementation, no process lifecycle management, no extra dependencies.

## Getting Started

### Claude Code Setup

1. Install Claude Code CLI: `npm install -g @anthropic-ai/claude-code`
2. Authenticate: `claude auth login` (sign in with your Claude Pro or Max subscription)
3. Verify it works: `claude -p "Hello"` should print a response

### Installation

#### Option 1: Pre-configured Build (Enterprise/RDS)

1. Build the solution in Release mode
2. Publish from Visual Studio (Right-click project > Publish)
3. Copy the publish folder to your deployment location
4. Run `Deploy\Install-OutlookAI.ps1` as Administrator:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
Unblock-File -Path "C:\OutlookAI\Install-OutlookAI.ps1"
cd C:\OutlookAI
.\Install-OutlookAI.ps1 -SourcePath "C:\OutlookAI"
```

#### Option 2: Per-User Install

1. Build and publish the solution
2. Run `setup.exe` from the publish folder
3. Open Outlook — the add-in is ready to use

### Building from Source

**Prerequisites:**

- Visual Studio 2022
- Office/SharePoint development workload
- .NET desktop development workload

**Build Steps:**

1. Clone this repository
2. Open `VSTO2\OutlookAI\OutlookAI.sln`
3. Restore NuGet packages
4. Build > Rebuild Solution

## Usage

1. Open Outlook and compose a new email (New, Reply, or Forward)
2. Click the **AI Assistant** button in the ribbon
3. The task pane opens on the right side

### Quick Actions

- Write your email draft first
- Click any Quick Action button (Proofread, Revise, etc.)
- Review the result and click **Insert**, **Replace**, or **Discard**

### Draft New Email

- Type your instructions (e.g., "Write a thank you email to John for the meeting")
- Click **Draft Email**
- Review and insert the result

### Custom Action

- Type any instruction in the Custom Action text box
- Click **Run Custom Action**
- The AI will apply your instruction to the current email content

## Configuration

Settings are stored in `%APPDATA%\OutlookAI\config.xml`.

Access the Settings panel by clicking the gear icon in the add-in. The default admin password is `admin`.

## Deployment Scripts

Located in the `Deploy` folder:

| Script | Purpose |
|---|---|
| `Install-OutlookAI.ps1` | Per-machine install for all users (RDS/Terminal Server) |
| `Uninstall-OutlookAI.ps1` | Remove the add-in |
| `Enable-OutlookAI-User.ps1` | Re-enable if Outlook disabled the add-in |

## Troubleshooting

<details>
<summary><strong>Add-in doesn't appear</strong></summary>

- Restart Outlook
- Check File > Options > Add-ins
- Run `Enable-OutlookAI-User.ps1`
</details>

<details>
<summary><strong>Add-in keeps getting disabled</strong></summary>

- Outlook's "Resiliency" feature may disable slow-loading add-ins
- Run `Enable-OutlookAI-User.ps1` or add it to logon scripts
</details>

<details>
<summary><strong>"Untrusted" or security errors</strong></summary>

- Ensure all files are unblocked (Right-click > Properties > Unblock)
- Or run: `Get-ChildItem -Path "C:\Program Files\OutlookAI" -Recurse | Unblock-File`
</details>

<details>
<summary><strong>"Claude Code CLI is not installed" error</strong></summary>

- Install Claude Code: `npm install -g @anthropic-ai/claude-code`
- Ensure `claude` is on your PATH (restart Outlook after installing)
</details>

<details>
<summary><strong>"Claude Code is not authenticated" error</strong></summary>

- Run `claude auth login` in a terminal and sign in with your Claude subscription
- Restart Outlook after authenticating
</details>

<details>
<summary><strong>Requests timing out</strong></summary>

- Check your internet connection
- Verify Claude Code works: `claude -p "Hello"`
- Claude may be temporarily overloaded — try again in a moment
</details>

## License

This project is a fork of [OutlookAI by kirklandsig](https://github.com/kirklandsig/OutlookAI), licensed under the [MIT License](LICENSE).

Original work: Copyright (c) 2026 kirklandsig

See the [LICENSE](LICENSE) file for the full license text.

## Acknowledgments

- [kirklandsig/OutlookAI](https://github.com/kirklandsig/OutlookAI) — Original project this fork is based on
- [Claude Code CLI](https://code.claude.com) — AI backend via Claude Code
