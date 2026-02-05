# Implementation Plan - Resolve M365 MCP Container Authentication

This plan addresses the issue where the `mcp-m365` container fails to authenticate or "freezes" due to a volume mounting conflict and TTY issues during the Device Login process.

## Problem Discovery
1.  **Volume Conflict**: `/root/.m365rc` is currently mounted as a **directory**, but the M365 Toolkit expects it to be a **file** for storing the auth token.
2.  **Interactive Hang**: The `server start` command triggers an interactive Device Login flow which hangs when run without a proper TTY or in non-interactive sessions.

## Proposed Changes

### 1. [Component: Docker Infrastructure]

#### [MODIFY] [docker-compose.yml](file:///c:/VBA/SGQ%201.65/docker-mcp/docker-compose.yml)
- Change the volume mount for `mcp-m365` to map to a directory instead of a specific file path that conflicts.
- **Change**: `- mcp-m365-data:/root/.m365-config`

#### [NEW] [entrypoint.sh](file:///c:/VBA/SGQ%201.65/docker-mcp/m365/entrypoint.sh)
- Add a script to handle the symlink between the persistent volume and the expected config file path.
- Logic: `ln -sf /root/.m365-config/.m365rc /root/.m365rc`

#### [MODIFY] [Dockerfile](file:///c:/VBA/SGQ%201.65/docker-mcp/m365/Dockerfile)
- Copy `entrypoint.sh` and set it as the `ENTRYPOINT`.

### 2. [Component: Scripts]

#### [MODIFY] [auth-m365.ps1](file:///c:/VBA/SGQ%201.65/scripts/auth-m365.ps1)
- Add a pre-check for the directory/file conflict.
- Enhance logging and instructions "Audit-grade".
- Ensure it uses `docker exec -it` for interactive login.
- Verify the existence of the token file after login.

## Verification Plan

### Automated Tests
- Run `docker exec mcp-m365 ls -l /root/.m365rc` to verify it's a symlink or a file (not a directory).
- Run `scripts/auth-m365.ps1` (manually by user) and verify success.

### Manual Verification
1.  **Initial Cleanup**:
    - `docker compose down`
    - `docker volume rm mcp-m365-data` (to clear the bad directory)
2.  **Deployment**:
    - `launch-docker-mcp.ps1`
3.  **Authentication**:
    - Execute `scripts/auth-m365.ps1`.
    - Follow the Device Login link and enter the code.
    - Confirm "Log-in Successful" in the console.
4.  **Functional Test**:
    - Verify that `m365agentstoolkit-mcp server start` runs without prompting for login after the first time.
