# Docker MCP Orchestration Walkthrough

This guide explains how to spin up the new Docker MCP environment including MarkItDown, GitHub, and DockerHub.

## 1. Environment Setup

First, navigate to the directory and set up your environment variables.

```powershell
cd "c:\VBA\SGQ 1.65\docker-mcp"
copy .env.example .env
```
> [!IMPORTANT]
> Edit `.env` and add your `GITHUB_PERSONAL_ACCESS_TOKEN`.

## 2. Launch Services

Build and start the containers. This will build the custom MarkItDown image and pull the others.

```powershell
docker-compose up -d --build
```

## 3. Verify Status

Check if all 6 containers are running:

```powershell
docker ps
```
You should see:
- `mcp-memory`
- `mcp-filesystem`
- `mcp-fetcher`
- `mcp-markitdown`
- `mcp-github`
- `mcp-dockerhub`

## 4. Connect Agent

Update your Client (Claude Desktop, Cursor, etc.) configuration with the JSON found in `README.md`. 
Specifically, add the entries for `docker-markitdown`, `docker-github`, and `docker-dockerhub`.

## 5. Test Tools

- **MarkItDown**: Ask the agent to "Convert c:/VBA/README.md to markdown using MarkItDown".
- **GitHub**: Ask the agent to "Search for 'mcp' on GitHub".
- **DockerHub**: Ask the agent to "Search for 'python' images on Docker Hub".
