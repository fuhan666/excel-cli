# Contributing to excel-cli

Thank you for your interest in contributing! This document defines the collaboration workflow for all contributors working on this repository.

## Workflow: GitHub Flow

We follow [GitHub Flow](https://docs.github.com/en/get-started/quickstart/github-flow):

1. Create a short-lived feature branch from `main`.
2. Make focused, atomic commits.
3. Open a Pull Request against `main`.
4. Ensure CI passes and address review feedback.
5. Merge via squash or merge commit once approved.

## Branch Naming Convention

All branches must use **lowercase kebab-case** with a type prefix:

| Prefix | Purpose |
|--------|---------|
| `feat/` | New features or enhancements |
| `fix/` | Bug fixes |
| `docs/` | Documentation-only changes |
| `refactor/` | Code restructuring without behavior changes |
| `test/` | Adding or updating tests |
| `chore/` | Maintenance, tooling, or dependency updates |

**Examples:**

- `feat/ai-inspect-commands`
- `fix/non-ascii-sheet-panic`
- `docs/update-install-guide`
- `chore/bump-dependencies`

## Commit Messages

All commit messages must follow [Conventional Commits](https://www.conventionalcommits.org/):

```
<type>(<optional-scope>): <description>

<optional-body>

<optional-footer>
```

Allowed types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`

**Examples:**

```
feat: add --peek command for headless range inspection
fix(parser): prevent panic on non-ascii sheet names
docs: install with --locked in README_zh
chore: bump version to 0.5.0
```

## Pull Request Requirements

### Language

**PR titles and PR descriptions must be written entirely in English.** This ensures consistency across the project and compatibility with automated tooling.

### Content

- Provide a clear summary of what changed and why.
- Reference related issues using `Closes #<number>` or `Relates to #<number>`.
- Ensure CI checks pass before requesting review.
