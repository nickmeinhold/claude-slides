# CLAUDE.md

Guidance for Claude Code when working on this repository.

## Project Overview

`claude-slides` is the TypeScript implementation of the `/slides` Claude Code slash command —
a Node.js CLI that generates Google Slides presentations from JSON configs.

The slash command itself (`slides.md`) and related skills (`/live-qa`) live in the sibling
repo [`nickmeinhold/claude-skills`](https://github.com/nickmeinhold/claude-skills). Skills there
invoke this build via:

```
npx --prefix "$CLAUDE_SLIDES_PATH" claude-slides --config <json>
```

This repo is intentionally separate from `claude-skills/` — when `~/.claude/commands` symlinks
to a directory containing `node_modules/`, Claude Code's command discovery walks every
`node_modules` README/CHANGELOG/LICENSE and exposes them as fake "skills," polluting the
system-reminder context (~6 k tokens per injection, ~60 k tokens / session in heavy use).

## Common Commands

```bash
# Build
npm run build

# Test
npm run test              # Run tests
npm run test:coverage     # Run with coverage (thresholds in vitest.config.ts)

# Develop
npm run dev -- --config example.json

# OAuth
npm run auth              # Google OAuth flow
```

## Architecture

```
src/
├── cli.ts              # Entry point, argument parsing
├── auth/
│   ├── oauth.ts        # OAuth2 flow with browser redirect
│   └── token-store.ts  # Token persistence (~/.claude-slides/tokens.json)
└── slides/
    ├── types.ts         # SlideConfig, SlideElement interfaces
    ├── config-loader.ts # Load JSON configs, interpolate variables
    ├── generator.ts     # Google Slides API calls
    └── templates.ts     # Color helpers, status emoji
```

### Key Concepts

- **EMU (English Metric Units)**: Google Slides uses EMU for positioning. Convert points: `points * 12700`
- **Slide dimensions**: Standard 16:9 is ~720 x 405 points
- **Color references**: Configs can use color names that resolve to RGB via theme.colors
- **Batch requests**: API limits ~100 requests per batch, code uses 50 for safety

### Authentication

OAuth tokens stored at `~/.claude-slides/tokens.json`.

**Setup (one-time):**
1. Add credentials to `.env` (in this repo, not committed):
   ```
   GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com
   GOOGLE_CLIENT_SECRET=your-client-secret
   ```
2. Run authentication:
   ```bash
   source .env
   export GOOGLE_CLIENT_ID GOOGLE_CLIENT_SECRET
   npm run auth
   ```

When invoked from the sibling `claude-skills` repo, the equivalent is:
```bash
npx --prefix "$CLAUDE_SLIDES_PATH" claude-slides --auth
```

**Note:** The CLI requires environment variables to be exported (not just present in `.env`).
After initial auth, tokens auto-refresh from `~/.claude-slides/`.

### Config Modes

1. **Static** (`--config`): Direct JSON with slide content
2. **Template** (`--template` + `--data`): JSON with `{{variables}}` interpolated from data file
3. **Legacy** (stdin): ReviewData JSON for PR review slides

## Testing Changes

After modifying CLI source:
```bash
npm run build
npx claude-slides --config test.json
# or, when invoked from a sibling skill:
npx --prefix "$CLAUDE_SLIDES_PATH" claude-slides --config test.json
```

Coverage thresholds are defined in `vitest.config.ts` and enforced by CI.

```bash
npm run test              # Quick test run
npm run test:coverage     # With coverage report
```

## File Conventions

- Source: `src/**/*.ts`
- Tests: `src/__tests__/*.test.ts`
- Build output: `dist/`
- Tokens/credentials: Not committed (in `.gitignore`)

## Environment Variables

The `.env` file (not committed) should contain Google OAuth credentials:

```bash
GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com
GOOGLE_CLIENT_SECRET=your-client-secret
```

`CLAUDE_SLIDES_PATH` (the path to a clone of this repo) lives in the sibling
`claude-skills` repo's `.env`, since that's where skills consume it.
