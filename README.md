# claude-slides

TypeScript implementation of the `/slides` Claude Code slash command.

This repo is intentionally separate from `claude-skills/` — when `~/.claude/commands` symlinks to a directory containing `node_modules/`, Claude Code's command discovery walks every `node_modules` README/CHANGELOG/LICENSE and exposes them as fake "skills," polluting the system-reminder context (~6 k tokens per injection, ~60 k tokens / session in heavy use).

The slash command itself (`slides.md`) lives in `claude-skills/`. It invokes this build via:

```
npx --prefix ~/git/individuals/nickmeinhold/claude-slides claude-slides --config <json>
```

## Build

```
npm install
npm run build      # produces dist/cli.js
```

## Develop

```
npm run dev        # tsx src/cli.ts
npm test           # vitest
```
