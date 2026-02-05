# Resolve M365 MCP Compatibility Issues

The Microsoft 365 Agents Toolkit MCP server is failing to execute correctly on the current Alpine-based Docker image (silent failure with exit code 1). This is likely due to library incompatibilities (musl vs glibc).

## Proposed Changes

### [Component] Docker Configuration

#### [MODIFY] [Dockerfile](file:///c:/VBA/SGQ%201.65/docker-mcp/m365/Dockerfile)
- Change base image from `node:18-alpine` to `node:18-slim`.

## Verification Plan

### Automated Steps
1. Rebuild the M365 container using the launch script:
   ```powershell
   .\launch-docker-mcp.ps1
   ```
2. Verify the binary can now output its version or help:
   ```powershell
   $docker = "C:\Program Files\Docker\Docker\resources\bin\docker.exe"
   & $docker exec mcp-m365 /usr/local/bin/m365agentstoolkit-mcp --help
   ```

### Manual Verification
1. Run the updated authentication script:
   ```powershell
   .\scripts\auth-m365.ps1
   ```
2. Confirm the device login code appears in the terminal.
3. User completes login at [microsoft.com/devicelogin](https://microsoft.com/devicelogin).
