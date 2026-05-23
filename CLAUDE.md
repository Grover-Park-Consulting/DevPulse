# DevPulse — Project Context for Claude Code

DevPulse is a Microsoft Access application for tracking AI-assisted development sessions. It records projects, steps, sessions, revision cycles, and tools used — measuring the thing it helps you do. This file gives Claude Code the context it needs to work effectively with this codebase.

## What Was Built Here

DevPulse was designed, built, and tested *entirely* using Claude Code. Every module, form, report, query, and SQL object was written in collaboration with an AI assistant. The sample data included with this distribution is the real build record — 21 tasks, 25 sessions, and the *revision cycle* counts that show exactly where the hard problems were.

## How It Was Built — The Development Workflow

Each task followed this sequence, without exception:

**Step 1 — Goals and Requirements**
Human and AI agree on what needs to be built before touching any code. Vague tasks produce poor results; the discipline of defining scope before starting compounds across the whole project.

**Step 2 — Written Plan**
A written proposal reviewed and refined before any implementation begins. Claude raises concerns; the developer decides what makes it into the build.

**Step 3 — Plan Review**
Pros and cons of each approach discussed openly. This is where problems get caught before they become revision cycles.

**Step 4 — Build**
Claude writes; the developer reviews, tests, and corrects. All Access object changes go through the live `.accdb` via MCP tools — VCS source files are reference copies, never edited directly.

**Step 5 — Test**
Each task tested in the live application before it is closed. Type checking and compilation verify code correctness; only live testing verifies feature correctness.

**Step 6 — Close**
Session metrics recorded: revision cycles, context resets, tools used, session timing. Key decisions and patterns captured while they are fresh. This is what makes it possible to resume the next day without reconstructing context.

**Step 7 — Memory**
What was built, what was tested, what passed, what is next — written to persistent memory immediately after each task closes, not at the end of the session.

## The Cowboy-to-Citizen Arc

The most honest observation from the human building DevPulse: the discipline AI-assisted development requires made us both better developers.

Early in the build, the instinct was to move fast — sketch an idea, start coding, adjust as we went. That is the Cowboy approach. It works when you are the only one who needs to understand the code, requirements are vaguely defined, and mistakes are cheap to fix.

Working with Claude Code changed that instinct. The benefits of working with AI compound with clear requirements:

- A well-defined task produces better code faster than a vague one.
- Context resets — when a conversation window fills and work must resume fresh — punish underspecified work because everything undocumented has to be reconstructed.
- The session-close protocol was not bureaucracy; it was the mechanism that made the next session possible without losing momentum.

An unanticipated, but highly appreciated, benefit of our more disciplined approach was that towards the end of active development, we simply didn't see the dreaded "Out of Context" warning during most sessions. We were focused, on task, with clear goals and more efficient use of our resources. It felt more relaxed even though we got more work done.

By the end of the build, we were writing task descriptions before writing code, capturing decisions as they were made, and using the revision cycle data to honestly assess where the AI added friction and where it saved it.

That is the Citizen approach: working in concert with a team, even when that team is one human and one AI.

## Codebase Structure

```text
DevPulseSSE.accdb.src/    VCS export — SQL Server edition front end
access_ace/
  DevPulseACE.accdb.src/  VCS export — ACE edition front end
  DevPulse_BE.accdb.src/  VCS export — Access back end
sql/
  DevPulse_Setup_SQLServer.sql  Schema + seed data + sample data (Tasks 1-21)
```

**Key modules:**

- `modAppSetup` — startup, linking, connection string, lookup cache
- `modSQLConnect` — first-run connection setup (SSE edition only)
- `modCodeLibraryExtracts` — runtime utilities (error handling, TempVars, lookup cache)
- `modRibbon` — ribbon event dispatch

**Key forms:**

- `frmMain` — dashboard; ribbon-driven navigation, live metrics, chart
- `frmProject` / `frmStep` / `frmSession` — the data entry hierarchy
- `frmSetupConnection` — first-run SQL Server credential setup (SSE edition)

## Tools Used to Build This

| Tool | Role |
| ------ | ------ |
| Claude Code | Primary AI assistant — all VBA, SQL, and Access structure |
| Access Explorer MCP | Live `.accdb` manipulation from within Claude Code |
| SQL MCP | Direct SQL Server queries during development |
| PowerShell | File generation, SQL execution, git operations |
| Access | Testing — every task verified in the live application |

## License

MIT. See [LICENSE](LICENSE).
