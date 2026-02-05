# Task: Resolve M365 MCP Container Authentication and Execution Issues

- [ ] Diagnostic of `docker exec` freezing [ / ]
    - [ ] Inspect existing configuration files (`Dockerfile`, `docker-compose.yml`, `mcp.json`)
    - [ ] Test container connectivity and basic command execution
- [ ] Implement robust Device Login process [ ]
    - [ ] Modify `auth-m365.ps1` for interactive login
    - [ ] Ensure token persistence in `/root/.m365rc`
- [ ] Optimize Docker configuration [ ]
    - [ ] Update `Dockerfile` for better stability
    - [ ] Refine `docker-compose.yml` if necessary
- [ ] Verification and "Audit-grade" finalization [ ]
    - [ ] Test end-to-end authentication and server start
    - [ ] Add robust error handling and logging
    - [ ] Update documentation (`docs/ARCHITECTURE_ET_PLAN.md`)
