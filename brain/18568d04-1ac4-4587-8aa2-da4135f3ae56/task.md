# Implementing Smart Import via MCP MarkItDown

- [x] Research & Design <!-- id: 19 -->
    - [x] List available MCP tools for `mcp-markitdown` <!-- id: 20 -->
    - [x] Research target VBA integration (Ribbon button + Module) <!-- id: 21 -->
    - [x] Define the data flow (File -> Markdown -> Parsing -> VBA) <!-- id: 22 -->
- [/] Implementation - Backend/PowerShell <!-- id: 23 -->
    - [x] Create a bridge script for AI-assisted parsing if necessary <!-- id: 24 -->
- [/] Implementation - VBA <!-- id: 25 -->
    - [x] Create `modSGQSmartImport.bas` <!-- id: 26 -->
    - [x] Implement `ImportDocumentViaMCP` procedure <!-- id: 27 -->
    - [x] Add Ribbon button to `customUI14.xml` and handle callback <!-- id: 28 -->
- [ ] Verification <!-- id: 29 -->
    - [ ] Test with a sample PDF/Docx <!-- id: 30 -->
    - [ ] Validate Data Sanitization (CPA standards) <!-- id: 31 -->
- [ ] Documentation <!-- id: 32 -->
    - [ ] Update `CHANGELOG.md` and `walkthrough.md` <!-- id: 33 -->
