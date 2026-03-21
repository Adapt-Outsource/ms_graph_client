# Contributing Guide

Thank you for contributing to this project.

This repository provides a reusable SharePoint library (`adapt_sharepoint`) with a production-oriented quality gate:

- linting (`ruff`)
- static type checking (`mypy`)
- unit tests (`pytest`)

All pull requests should keep these checks green.

## Development Setup

1. Clone repository.
2. Install dependencies with uv:

```bash
uv sync
```

3. Optional: activate venv.

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

## Project Structure

```text
src/adapt_sharepoint/   # package source code
tests/                  # unit tests
README.md               # user documentation
CONTRIBUTING.md         # contributor instructions
```

## Development Workflow

1. Create a feature branch from `main`.
2. Implement changes in `src/adapt_sharepoint`.
3. Add or update tests in `tests/`.
4. Run quality checks locally.
5. Update docs (`README.md`) if API behavior changes.
6. Open PR with summary and test evidence.

## Quality Checks (Required)

Run all commands before opening PR.

```bash
uv run --with ruff ruff check .
uv run --with mypy mypy src tests
uv run --with pytest --with pytest-asyncio pytest -q
```

## Unit Test Guidelines

### Test Location

- Put tests under `tests/`.
- File naming: `test_<feature>.py`.
- Test naming: `test_<behavior>_<expected_result>`.

### What To Test

When adding/changing behavior, include tests for:

1. Happy path behavior.
2. Edge cases and invalid input.
3. Error handling paths.
4. Backward compatibility for existing public API behavior.

### Async Tests

This project uses async APIs. Write async tests like:

```python
async def test_example() -> None:
    ...
```

`pytest-asyncio` is already configured in `pyproject.toml`.

### Mocking External Calls

Do not call real Microsoft Graph in unit tests.

- Use fake clients/responses (as done in existing tests).
- Assert request path/headers/content where relevant.
- Keep tests deterministic and offline.

### Path Handling Tests

For file operations, test both:

1. `preserve_path=False` (default flat download behavior)
2. `preserve_path=True` (preserved folder structure)

Also test path-like inputs where possible (`str`, `Path`, custom `PathLike`).

## Coding Standards

- Python 3.11+
- Type hints required for new/modified code.
- Keep public API stable unless change is intentional and documented.
- Keep functions focused and side effects explicit.
- Avoid breaking existing behavior without tests and README updates.

## Documentation Requirements

Update `README.md` when any of the following changes:

1. Public method signature.
2. Parameter semantics/defaults.
3. Authentication behavior.
4. Runtime behavior (download/upload path behavior, etc.).

## Commit Guidance

### Commit Messages

Use clear, action-oriented commit messages, for example:

- `feat: add preserve_path flag to download API`
- `fix: support path-like objects in upload/download`
- `docs: update README download behavior examples`
- `test: add regression for path-like output_dir`

## Security and Secrets

Never commit:

- access tokens
- credentials
- private URLs with embedded secrets

Use environment variables or secure secret stores for sensitive values.

## Release Notes Suggestion

For user-visible changes, include a short note in PR description:

1. What changed.
2. Why it changed.
3. Any migration action needed by consumers.
