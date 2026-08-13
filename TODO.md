# TODO

- [ ] **Update the MCP server SDK — `ModelContextProtocol` 1.4.1 → latest.** The pin in `McpServer/OutlookAI.McpServer/OutlookAI.McpServer.csproj` carries the comment *"1.4.1 = latest stable on nuget.org as of 2026-07-23 (2.0.0 is still preview)"*. Re-checked against nuget.org on **2026-08-14**: 2.0.0, 2.1.0 and 2.2.0 have all shipped stable, so that claim is no longer true and the pin is three releases behind.
  - [ ] Review the 1.4 → 2.x breaking changes before bumping. Specifically: do `AddMcpServer(…).WithStdioServerTransport().WithToolsFromAssembly()` (`Program.cs`) and the `[McpServerToolType]` / `[McpServerTool]` attributes (`Tools/OutlookTools.cs`) survive unchanged, or did they move — e.g. into `ModelContextProtocol.Core`?
  - [ ] Check which MCP protocol revision 2.x negotiates. `ServerMetadata.cs` and the T3 suite currently pin `2025-06-18`; if the SDK's default moved, decide whether to follow it or hold the revision explicitly, and update the T3 assertion either way.
  - [ ] Re-run T1 (42 unit), T2 (42 live-Outlook) and T3 (20 stdio-conformance). T3 speaks raw JSON-RPC and is deliberately independent of the package on both ends — it is the honest check that wire behaviour did not change.
  - [ ] Verify the 21-tool surface and the `initialize` instructions string are unchanged, since the T1 tests pin both.
  - [ ] Re-stamp the csproj comment with a fresh `as of <date>`, or delete the pin rationale if the version is no longer being held back deliberately.

- [ ] **Retire v3 planning ignores** — once the local v3 planning files (`v3.MD`, `Docs/v3-probes/`) are no longer needed:
  - [ ] remove the "v3 planning documents" section at the bottom of `.gitignore`
  - [ ] delete the local plan-doc backup folder (location documented in v3.MD §0.8 D16 on the machine that holds it)
  - [ ] delete this TODO entry (and this file if empty)
