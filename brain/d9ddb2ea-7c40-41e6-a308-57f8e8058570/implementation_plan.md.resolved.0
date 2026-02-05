# Implementation Plan - Docker MCP Orchestration

This plan outlines the steps to configure and orchestrate MarkItDown, GitHub, and DockerHub MCP servers using Docker, ensuring a secure and reproducible environment.

## User Review Required
> [!IMPORTANT]
> **GitHub Token**: The GitHub MCP requires a `GITHUB_PERSONAL_ACCESS_TOKEN`. This should be set in the `.env` file (not committed to git).
> **DockerHub Access**: The DockerHub MCP (read-only) might require credentials or just operate on public repos. I will assume public access for now unless valid credentials are provided in `.env`.

## Proposed Changes

### Configuration
#### [MODIFY] [docker-compose.yml](file:///c:/VBA/SGQ%201.65/docker-mcp/docker-compose.yml)
- Add `mcp-markitdown` service (Build from context).
- Add `mcp-github` service (Node image).
- Add `mcp-dockerhub` service (Node image or similar - TBD based on exact package).
- Ensure all services join `mcp-network`.

#### [NEW] [docker-mcp/markitdown/Dockerfile](file:///c:/VBA/SGQ%201.65/docker-mcp/markitdown/Dockerfile)
- Python 3.10-slim base.
- Install `markitdown` and `mcp` (python sdk).
- Entrypoint to run the MCP server.

#### [NEW] [docker-mcp/markitdown/server.py](file:///c:/VBA/SGQ%201.65/docker-mcp/markitdown/server.py)
- A simple Python script to wrap `markitdown` library as an MCP server (if a CLI server isn't available out of the box).
- *Self-correction*: I will check if `markitdown` has a CLI that speaks MCP. If not, I'll write a wrapper.

#### [MODIFY] [README.md](file:///c:/VBA/SGQ%201.65/docker-mcp/README.md)
- Update "Services Inclus" section.
- Add configuration examples for the new MCPs in `claude_desktop_config.json`.

#### [NEW] [.env.example](file:///c:/VBA/SGQ%201.65/docker-mcp/.env.example)
- Template for environment variables (GITHUB_TOKEN, etc.).

## Verification Plan

### Automated Tests
- Run `docker-compose config` to validate syntax.
- Run `docker-compose up -d --build` to verify containers start.
- Test basic functionality via `docker exec`:
    - **MarkItDown**: `docker exec -i mcp-markitdown python server.py < test.file` (simulate stdio).
    - **GitHub**: `docker exec mcp-github npx @modelcontextprotocol/server-github --version`.
    - **DockerHub**: Verify container health.

### Manual Verification
- Ask user to start the stack.
- Ask user to verify they can see the new servers in `docker ps`.
