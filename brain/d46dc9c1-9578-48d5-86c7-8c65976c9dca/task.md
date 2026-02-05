# Task: Debug M365 MCP Container

- [x] Analyze current configuration and authentication scripts
    - [x] Read `scripts/auth-m365.ps1`
    - [x] Read `mcp.json`
    - [x] Read `docker-mcp/m365/Dockerfile`
    - [x] Read `docker-mcp/docker-compose.yml`
- [ ] Investigate toolkit behavior in container
    - [ ] Attempt `server start` manually in the container
    - [ ] Capture errors related to "Device login"
- [ ] Implement fix for authentication/startup
    - [ ] Update `Dockerfile` or `docker-compose.yml` if environment variables or TTY are missing
    - [ ] Fix `auth-m365.ps1` if it points to the wrong commands
- [ ] Verify fix
    - [ ] Run health checks on `mcp-m365`
    - [ ] Confirm tools are accessible via MCP
