# Skills — Published

Public skills repo. Contains skills authored by Jonny that are published for others to use.

## What lives here

Only published/public skills: currently just `md-to-docx-template`.

Private skills live in `~/Dev/hub/` — each folder with a `SKILL.md` is a skill. Tracking (MANIFEST.md, USAGE-LOG.md) also lives at the hub root.

## Architecture

Skills from all sources are symlinked into `~/.claude/skills/` — this is what Claude Code reads at session start.

| Source | Location | What belongs here |
|--------|----------|-------------------|
| **Published** | `~/Dev/skills/skills/` | Skills published for others (this repo) |
| **Private** | `~/Dev/hub/` | Jonny's personal skills (folders with SKILL.md), tracking |
| **Third-party** | `~/.claude/third-party-skills/` | Cloned repos (Lenny, Vercel) |

## Adding a Published Skill

1. Create `~/Dev/skills/skills/<skill-name>/SKILL.md`
2. Install: `ln -s ~/Dev/skills/skills/<skill-name> ~/.claude/skills/<skill-name>`
3. Add a row to `~/Dev/hub/MANIFEST.md`

## Related Workspaces

- **Skill tracking:** `~/Dev/hub/MANIFEST.md` and `~/Dev/hub/USAGE-LOG.md`
- **Automations:** `~/Dev/hub/` (see its CLAUDE.md for automation docs)
- **Dev index:** `~/Dev/CLAUDE.md`
