# Skills

Agent skills library — the central hub for managing Claude Code skills.

## Architecture

Skills live in three source locations. All are installed via symlink into `~/.claude/skills/`.

| Source | Location | What belongs here |
|--------|----------|-------------------|
| **Personal** | `~/Dev/skills/skills/` | Custom skills authored by Jonny |
| **Automations** | `~/Dev/jonny-automations/skills/` | Skills embedded in the automations repo |
| **Third-party** | `~/.claude/third-party-skills/` | Cloned repos (Lenny, Vercel) |

### Install directory

All skills are symlinked into `~/.claude/skills/` — this is what Claude Code reads at session start.

```
~/.claude/skills/<skill-name> → <source-path>
```

## Key Files

| File | Purpose |
|------|---------|
| `MANIFEST.md` | Full index of all skills with status ratings (Fave/Active/Tried/Unused) |
| `USAGE-LOG.md` | Append-only log of skill usage, updated by skill-tracker |
| `skills/skill-tracker/SKILL.md` | Meta-skill that auto-logs usage and promotes Unused → Tried |

## Adding a New Personal Skill

1. Create `~/Dev/skills/skills/<skill-name>/SKILL.md`
2. Install: `ln -s ~/Dev/skills/skills/<skill-name> ~/.claude/skills/<skill-name>`
3. Add a row to `MANIFEST.md`

## Installing a Third-Party Skill

1. Clone the repo into `~/.claude/third-party-skills/`
2. Symlink individual skills: `ln -s ~/.claude/third-party-skills/<repo>/<skill-path> ~/.claude/skills/<skill-name>`
3. Add rows to `MANIFEST.md` with ❌ Unused status

## Skill Tracker

The `skill-tracker` meta-skill runs silently after every skill invocation:
1. Appends an entry to `USAGE-LOG.md`
2. If the skill is ❌ Unused in the manifest, promotes it to 🔸 Tried

It is not user-invocable — Claude calls it internally.

## Related Workspaces

- **Business hub:** `/Users/Jonny/My Drive (jonny@humventures.com.au)/`
- **Automations:** `~/Dev/jonny-automations/` (see its CLAUDE.md for automation docs)
- **Dev index:** `~/Dev/CLAUDE.md`
