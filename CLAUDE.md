# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

AutoOfficeWork V2 is a Claude-powered office automation engine with a strict security boundary: **Claude plans (no real data), Python engine executes (with real data)**. Claude analyzes empty templates + natural language process descriptions to produce `execution_plan.json`, which the on-premise engine runs against actual data.

## Commands

```bash
# Install (editable dev mode)
pip install -e ".[dev]"

# Run an execution plan
autooffice run execution_plan.json --data ./data/ [-v]

# Validate a plan (dry-run, no execution)
autooffice validate execution_plan.json [-v]

# Cache management
autooffice cache list
autooffice cache run <plan_id> --data ./data/

# Run all tests
pytest

# Run a single test file
pytest tests/test_engine/test_runner.py

# Run a specific test
pytest tests/test_engine/test_actions/test_excel_actions.py::test_open_file -v

# Coverage
pytest --cov=autooffice
```

## Architecture

### Two-Phase Separation

```
Phase 1: Claude (no data)              Phase 2: Python Engine (with data)
├─ Analyze empty template structure    ├─ Parse execution_plan.json (Pydantic)
├─ 5-phase thinking (SKILLS.md)        ├─ PlanRunner executes steps sequentially
├─ Produce execution_plan.json         ├─ ActionHandlers do atomic operations
└─ Produce thinking_log.txt            └─ Produce ExecutionLog with StepResults
```

### Core Execution Flow

`ExecutionPlan` → `PlanRunner.run()` → for each `Step`: resolve `$variable` params via `EngineContext` → lookup `ActionHandler` in registry → execute → check `expect` conditions → handle `on_fail` strategy → store results via `store_as`.

### Key Modules

- **`src/autooffice/models/`** — Pydantic models: `ExecutionPlan` (plan schema with Steps, ActionType enum, OnFailAction), `ActionResult`/`StepResult`/`ExecutionLog` (results)
- **`src/autooffice/engine/runner.py`** — `PlanRunner`: sequential step execution, dry-run validation, expect checking
- **`src/autooffice/engine/context.py`** — `EngineContext`: runtime state (open workbooks, variables dict for `$ref` resolution, file paths)
- **`src/autooffice/engine/actions/`** — Registry-based handler pattern. Each handler extends `ActionHandler` ABC. New actions: create handler class → register in `build_default_registry()` in `__init__.py`
- **`src/autooffice/cache/`** — `PlanCache`: template-hash + task-type based plan reuse, stored at `~/.autooffice/cache/`
- **`src/autooffice/cli.py`** — Click CLI entry point
- **`schemas/execution_plan.schema.json`** — JSON Schema contract between Claude and engine
- **`skills/SKILLS.md`** — Claude's 5-phase thinking process instructions for plan generation

### Action Handler Registry Pattern

All 11 action types (OPEN_FILE, READ_COLUMNS, READ_RANGE, WRITE_DATA, CLEAR_RANGE, RECALCULATE, SAVE_FILE, VALIDATE, FORMAT_MESSAGE, SEND_MESSENGER, LOG) are registered via `build_default_registry()`. To add a new action:
1. Create a handler class extending `ActionHandler` in `engine/actions/`
2. Implement `execute(params, ctx) → ActionResult`
3. Register with `ActionType.XXX.value` key in `build_default_registry()`

### Inter-Step Data Flow

Steps communicate via `EngineContext.variables`. A step with `store_as: "my_data"` saves its result; later steps reference it as `$my_data` in params. `ctx.resolve_params()` handles substitution.

### Failure Handling

Each step has `on_fail`: `STOP` (halt), `SKIP` (ignore), `RETRY` (not yet implemented), `WARN_AND_CONTINUE`. The `expect` field defines success criteria checked after execution.

## Conventions

- Language: Korean comments/docs, English code identifiers
- Python ≥ 3.11, Pydantic v2 for all models
- xlwings for Excel manipulation (requires Excel installed)
- Steps are 1-indexed and must be continuous (validated by Pydantic)
- All file operations go through `EngineContext` which tracks workbooks and paths
- Test fixtures in `tests/conftest.py` provide `engine_ctx`, `runner`, sample Excel files

## Stubs / Not Yet Implemented

- `SEND_MESSENGER`: needs sidecar API integration
- `SEND_EMAIL`, `GENERATE_PPT`: schema-defined but no handler
- `RETRY` on_fail: schema-defined but logic not implemented
- `require_confirm`: logged but no interactive prompt
