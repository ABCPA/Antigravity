# Implementation Plan - Dockerize GitHub MCP Server

The goal is to integrate the GitHub MCP server into the existing Docker environment, ensuring it runs reliably and can be accessed by the agent/client.

## User Review Required

> [!IMPORTANT]
> **GitHub Personal Access Token Required**
> You will need to generate a GitHub Personal Access Token (Classic) with `repo` permissions and add it to your `.env` file as `GITHUB_PERSONAL_ACCESS_TOKEN`.

## Proposed Changes

### Docker Configuration

#### [NEW] [Dockerfile](file:///c:/VBA/SGQ 1.65/docker-mcp/github/Dockerfile)
Create a Dockerfile to pre-install `@modelcontextprotocol/server-github`.
- Base: `node:18-alpine`
- Install: `npm install -g @modelcontextprotocol/server-github`
- Entrypoint: `tail -f /dev/null` (keeps container alive for `docker exec`)

#### [MODIFY] [docker-compose.yml](file:///c:/VBA/SGQ 1.65/docker-mcp/docker-compose.yml)
- Update `mcp-github` service to use `build: ./github` instead of just `image`.
- Ensure `GITHUB_PERSONAL_ACCESS_TOKEN` is passed to the container.

### Environment

#### [MODIFY] [.env](file:///c:/VBA/SGQ 1.65/.env)
- Add `GITHUB_PERSONAL_ACCESS_TOKEN=` placeholder.

### MCP Configuration

#### [MODIFY] [mcp.json](file:///c:/VBA/SGQ 1.65/mcp.json)
- Add `github` server configuration.
- Use `docker exec -i mcp-github npx @modelcontextprotocol/server-github` (or direct binary if in path).

## Verification Plan

### Manual Verification
1.  **Build and Start**: Run `docker-compose up -d --build` in `docker-mcp`.
2.  **Check Logs**: Verify `docker logs mcp-github` shows it's running (waiting).
3.  **Test Execution**: Run `docker exec -i mcp-github npx @modelcontextprotocol/server-github` manually to see if it starts (it should wait for stdin).
4.  **MCP Integration**: Restart the Agent/VS Code and ask it to "List my GitHub repositories" to verify connectivity (requires valid token).
