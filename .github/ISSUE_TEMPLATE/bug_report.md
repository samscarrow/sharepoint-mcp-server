---
name: Bug report
about: Create a report to help us improve
title: '[BUG] '
labels: bug
assignees: ''

---

**Describe the bug**
A clear and concise description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:
1. Configure environment with '...'
2. Run command '....'
3. Call tool '....'
4. See error

**Expected behavior**
A clear and concise description of what you expected to happen.

**Error Output**
If applicable, add error messages or logs to help explain your problem.

```
Paste error output here
```

**Environment (please complete the following information):**
 - OS: [e.g. Windows 11, macOS 14, Ubuntu 22.04]
 - Node.js version: [e.g. 18.17.0]
 - SharePoint MCP Server version: [e.g. 0.1.0]
 - MCP Client: [e.g. Claude Desktop, MCP Inspector]

**SharePoint Configuration:**
 - SharePoint version: [e.g. SharePoint Online, SharePoint 2019]
 - Azure app permissions: [list the permissions granted]
 - Tenant type: [e.g. Commercial, GCC, GCC High]

**Additional context**
Add any other context about the problem here.

**Configuration (remove sensitive information)**
```json
{
  "mcpServers": {
    "sharepoint-mcp-server": {
      "command": "node",
      "args": ["/path/to/sharepoint-mcp-server/build/index.js"],
      "env": {
        "SHAREPOINT_URL": "https://yourtenant.sharepoint.com",
        "TENANT_ID": "redacted",
        "CLIENT_ID": "redacted",
        "CLIENT_SECRET": "redacted"
      }
    }
  }
}
