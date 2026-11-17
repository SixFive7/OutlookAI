/**
 * @name MCP tools whose parameters are untrusted input
 * @description The number of methods that carry ModelContextProtocol's [McpServerTool] attribute
 *              and have a parameter the active threat models count as untrusted input.
 *              Tools/Invoke-CodeQL.ps1 requires it to equal the number of tools the source
 *              declares. A model that stops matching then fails the run: an SDK rename, or a
 *              package build-mode none could not restore.
 * @kind metric
 * @tags summary
 * @id outlookai/mcp-tools-with-untrusted-parameters
 */

import csharp
import semmle.code.csharp.security.dataflow.flowsources.FlowSources

select count(Method tool |
    tool.fromSource() and
    tool.getAnAttribute()
        .getType()
        .hasFullyQualifiedName("ModelContextProtocol.Server", "McpServerToolAttribute") and
    exists(ActiveThreatModelSource source | source.asParameter() = tool.getAParameter())
  )
