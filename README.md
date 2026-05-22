# DevPulse

**A greenfield Microsoft Access project built to answer one question: what happens when a developer steps back from the keyboard, takes the role of Project Manager, and lets an AI assistant do 100% of the development work?**

DevPulse is the answer. It is a session-tracking application for AI-assisted development projects — and it was built entirely by Claude Code, Anthropic's AI development CLI, under the direction of a human project manager who wrote no code.

Every requirement was defined by a person. Every line of VBA, SQL, and Access structure was written by the AI. Every task was reviewed, approved, tested, and accepted by the person. The AI handled the implementation; the person handled the judgment.

## What DevPulse Does

DevPulse tracks AI-assisted development sessions at three levels:

- **Project** — the initiative being built
- **Step** — a discrete task within the project (design, build, test, refactor)
- **Session** — a working session on a step, recording start/end time, revision cycles, context resets, and tools used

The headline metric is **revision cycles** — how many times something had to be tried again before it worked. That number, aggregated across steps and projects, is an honest measure of how much back-and-forth AI-assisted development actually takes.

## The Meta Angle

The sample data included with DevPulse is the real build record of DevPulse itself — 21 tasks, 25 sessions, and the revision cycle counts from the original build. The app was tracked in itself while being built. Task 16 (Ribbon XML and startup relink) took 11 revision cycles. Most tasks took 0–3. The data is honest.

## Two Backend Options

DevPulse ships with two front ends, each paired with a different backend:

| Edition | Backend | Best for |
|---------|---------|----------|
| **SQL Server Express** | SQL Server (free) | Teams, shared databases, distributed access |
| **Access BE** | Access .accdb | Single-user, fully self-contained, no server required |

Both editions include the same sample data and functionality. The SQL Server edition is the primary offering.

## Getting Started

> Full setup instructions are in progress and will be added here shortly.

**SQL Server edition:** Run `sql/DevPulse_Setup_SQLServer.sql` against a SQL Server or SQL Server Express instance, then open `DevPulseSSE.accdb` and enter your connection details when prompted.

**Access BE edition:** Open `DevPulseACE.accdb` — it connects automatically to the included `DevPulse_BE.accdb`.

Both `.accdb` files are distributed as GitHub Releases — see the [Releases](../../releases) page.

## How It Was Built

See [CLAUDE.md](CLAUDE.md) for the full story — the development methodology, the tools used, and the shift from writing code to managing a project.

## License

MIT — see [LICENSE](LICENSE).
