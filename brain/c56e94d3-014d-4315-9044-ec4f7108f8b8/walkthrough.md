# Walkthrough - Dockerized MCP Environment

I have successfully created and optimized the "Audit-grade" Docker environment for your MCP servers.

## 📂 Deliverables

The following files were created in `c:\VBA\SGQ 1.65\docker-mcp\`:

1.  **[launch_env.ps1](file:///c:/VBA/SGQ%201.65/docker-mcp/launch_env.ps1)**
    *   **Automated Launcher**: Detects Docker (even if not in PATH), fixes environment, and launches containers.
    *   **Usage**: Right-click > "Run with PowerShell".

2.  **[docker-compose.yml](file:///c:/VBA/SGQ%201.65/docker-mcp/docker-compose.yml)**
    *   **Architecture**: "Keep-Alive" containers.
    *   **Reason**: Prevents file locking issues. The containers run `tail -f /dev/null`, acting as "toolboxes".
    *   **Usage**: Your Agent spawns the specific tool it needs inside the container via `docker exec`.

3.  **[.env](file:///c:/VBA/SGQ%201.65/docker-mcp/.env)**
    *   Environment configuration file.

4.  **[README.md](file:///c:/VBA/SGQ%201.65/docker-mcp/README.md)**
    *   Complete guide on how to start the environment and connect your AI agent.
    *   **Updated**: Now contains the precise `docker exec` JSON configuration for your Agent.

## 🚀 How to Launch

**Option 1 (Recommended)**:
Run the script **[launch_env.ps1](file:///c:/VBA/SGQ%201.65/docker-mcp/launch_env.ps1)**.

**Option 2 (Manual)**:
```powershell
cd "c:\VBA\SGQ 1.65\docker-mcp"
docker compose up -d --force-recreate
```

## 🔌 Next Step: Agent Connection

Copy the JSON configuration from **[README.md](file:///c:/VBA/SGQ%201.65/docker-mcp/README.md)** into your `claude_desktop_config.json` or Agent settings to enable the tools.
