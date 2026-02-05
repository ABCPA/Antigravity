# Implementation Plan - Phase 41: Memory Initialization

## Goal
Populate the newly created **Memory MCP (Knowledge Graph)** with the critical "Business Physics" of the SGQ 1.65 project.
This ensures the Agent "knows" the rules rather than "reading" them every time.

## Strategy
Since we cannot write to the Memory MCP directly from the host (it requires an MCP client), we will:
1.  **Extract** knowledge into a structured "Seed File" (`docs/KNOWLEDGE_SEED.md`).
2.  **Prompt** the User's Agent (Claude) to read this file and execute the `create_entity` / `create_relation` tools.

## Content of `KNOWLEDGE_SEED.md`
We will structure the data into **Entities** and **Relations**:

### 1. Global Rules (Entities)
*   `Rule: CP1252_Encoding`: "All .bas/.cls files MUST be Windows-1252."
*   `Rule: 4_Layer_Architecture`: "UI -> Services -> Core -> DevTools."
*   `Rule: Audit_Grade`: "No On Error Resume Next without Try* wrapper."

### 2. Architecture (Entities & Relations)
*   `Layer: UI` *contains* `Module: modRibbonSGQ`
*   `Layer: Core` *contains* `Module: modSGQUtilitaires`
*   `Module: modAppStateGuard` *is_critical_for* `Operation: StateManagement`

## Verification
*   **Manual**: User asks Agent "What are the encoding rules?" -> Agent queries Memory -> Answers "CP1252".
