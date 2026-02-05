# Task: Debug GitHub MCP Authentication

- [x] Investigate Configuration <!-- id: 0 -->
    - [x] Read `.env` file (verify variable names) <!-- id: 1 -->
    - [x] Read `docker-compose.yml` and `Dockerfile` <!-- id: 2 -->
    - [x] Analyze invalid Authorization header error <!-- id: 3 -->
- [x] Fix Authentication Issue <!-- id: 4 -->
    - [x] Request user to provide `GITHUB_COPILOT_TOKEN` or ignore if unused <!-- id: 6 -->
- [x] Verify Connection <!-- id: 5 -->
    - [x] Check if `docker` is in PATH <!-- id: 7 -->
    - [x] Verify `mcp-github` container is running <!-- id: 8 -->
    - [x] Fix `mcp.json` to use absolute path for docker <!-- id: 9 -->
