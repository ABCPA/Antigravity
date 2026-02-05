# Task: Resolve M365 MCP Container Authentication and Execution Issues

- [x] Diagnostic of `docker exec` freezing
    - [x] Inspect existing configuration files (`Dockerfile`, `docker-compose.yml`, `mcp.json`)
    - [x] Test container connectivity and basic command execution
- [x] Implement robust Device Login process
    - [x] Modify `auth-m365.ps1` for interactive login
    - [x] Ensure token persistence in `/root/.m365rc`
- [x] Optimize Docker configuration
    - [x] Update `Dockerfile` for better stability
    - [x] Refine `docker-compose.yml` if necessary
- [x] Verification and "Audit-grade" finalization
    - [x] Test end-to-end authentication and server start
    - [x] Add robust error handling and logging
    - [x] Update documentation (`docs/ARCHITECTURE_ET_PLAN.md`)
