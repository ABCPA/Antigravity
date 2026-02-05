# Plan: Fix M365 MCP Container Authentication

The objective is to ensure the `mcp-m365` container starts correctly and supports the Microsoft 365 authentication flow (Device Login).

## Proposed Changes

### Configuration
#### [MODIFY] [mcp.json](file:///C:/VBA/SGQ%201.65/mcp.json)
- Verify and update the command used to start the M365 MCP server.

### Docker
#### [MODIFY] [Dockerfile](file:///C:/VBA/SGQ%201.65/docker-mcp/m365/Dockerfile)
- Ensure all necessary environment variables for authentication are present.
- Fix any path issues for the toolkit binary.

### Scripts
#### [MODIFY] [auth-m365.ps1](file:///C:/VBA/SGQ%201.65/scripts/auth-m365.ps1)
- Align the script with the actual CLI commands of the toolkit.

## Verification Plan
### Manual Verification
- Execute the authentication command from the host using `docker exec`.
- Verify the container logs for successful startup signal.
- List tools from the MCP server to confirm it is operational.
