# DevPulse

**A greenfield Microsoft Access project built to answer one question: what happens when a developer steps back from the keyboard, takes the role of Project Manager, and lets an AI assistant do 100% of the development work?**

DevPulse is the answer. It is a session-tracking application for AI-assisted development projects — and it was built entirely by Claude Code, Anthropic's AI development CLI, under the direction of a human project manager who wrote no code.

Every requirement was defined by a person. Every line of VBA, SQL, and Access structure was written by the AI. Every task was reviewed, approved, tested, and accepted by the person. The AI handled the implementation; the person handled the judgment.

## What DevPulse Does

DevPulse tracks AI-assisted development sessions at three levels:

- **Project** — the initiative being built
- **Step** — a discrete task within the project (design, build, test, refactor)
- **Session** — a working session on a step, recording start/end time, revision cycles, context resets, and tools used

### Revision Cycles — the Magic Number in DevPulse

The headline metric is **revision cycles** — how many times something had to be tried again before it worked. That number, aggregated across steps and projects, is an honest measure of how much back-and-forth AI-assisted development actually takes.

## The Meta Angle

The sample data included with DevPulse is the real build record of DevPulse itself — 21 tasks, 25 sessions, and the revision cycle counts from the original build. The app was tracked in itself while being built. Task 16 (Ribbon XML and startup relink) took 11 revision cycles. Most tasks took 0–3. The data is honest.

What does revision cycles tell us? It's the answer to a complaint you hear all the time. "AI makes mistakes." The DevPulse project revealed that to be true. It also revealed that it doesn't matter nearly as much as you might think.

Task 16 was building the Ribbon XML from the human developer's description. That's something the AI, Claude Code in this case, didn't know much about. Moreover, documentation on the Office Ribbon is sparse. That led to a rapid cycle of trial and error in which Claude Code rapidly found and discarded 10 ways that didn't work, mostly without intervention by the human PM. That amounts to a heuristic problem solving approach implemented by Claude Code, and that's something you hear a lot less about.

DevPulse surfaced that capability over and over again. Sure, an AI makes mistakes, but it also knows how to recover from them.

## Watch the Build

- **[DevPulse Starter — Demo Walkthrough](https://youtu.be/vzmxFtlfVAE)** — A tour of the application: dashboard, data entry, reports, and the session metrics that make it useful.
- **[PM Lessons Learned](https://youtu.be/dC-TJrjU7hA)** — Three ways the build data changed how the project was managed. For developers who want to understand what AI-assisted development actually looks like in practice.
- **[Tool Stack Walkthrough](https://youtu.be/bprKDNBuLPw?si=5oacKEcgg_2Ohx9A)** — The exact tools used to build DevPulse: Claude Code, Access Explorer MCP, VS Code, Git, GitHub, JoyfullService VCS, PowerShell, SSMS, and DevPulse itself. A 44-minute field report.

## Two Backend Options

DevPulse ships with two front ends, each paired with a different backend:

| Edition                | Backend           | Best for                                              |
| ---------------------- | ----------------- | ----------------------------------------------------- |
| **SQL Server Express** | SQL Server (free) | Teams, shared databases, distributed access           |
| **Access BE**          | Access .accdb     | Single-user, fully self-contained, no server required |

Both editions include the same sample data and functionality. The SQL Server edition is the primary offering.

## Getting Started

### Prerequisites

Both editions require **Microsoft Access 2016 or later** (including Microsoft 365).

Download the `.accdb` files from the [Releases](../../releases) page — binary files are not stored in the repository.

---

### SQL Server Edition

**Additional requirement:** A SQL Server or [SQL Server Express](https://www.microsoft.com/en-us/sql-server/sql-server-downloads) instance. SQL Server Express is free and sufficient for personal use.

1. Download `DevPulseSSE.accdb` from the [Releases](../../releases) page.
2. Create a new `DevPulse` database on your SQL Server instance, then run `sql/DevPulse_Setup_SQLServer.sql` against it. The script creates all tables, views, seed data, and the sample build history.
   - **SSMS:** Create the database via Object Explorer, open the script, select `DevPulse` as the target, and click **Execute**.
   - **Command line:** `sqlcmd -S <server> -Q "CREATE DATABASE DevPulse"` then `sqlcmd -S <server> -d DevPulse -i DevPulse_Setup_SQLServer.sql`
3. Create a SQL Server login for DevPulse. DevPulse uses SQL Server Authentication — not Windows Authentication. Choose any login name and password you like; the example below uses `DevPulseUser`. Run this against your instance:

   ```sql
   CREATE LOGIN DevPulseUser WITH PASSWORD = 'YourPassword';
   USE DevPulse;
   CREATE USER DevPulseUser FOR LOGIN DevPulseUser;
   ALTER ROLE db_datareader ADD MEMBER DevPulseUser;
   ALTER ROLE db_datawriter ADD MEMBER DevPulseUser;
   ```

   > **SQL Server Express note:** If you see a login failure, your instance may be configured for Windows Authentication only. In SSMS, right-click the server → **Properties** → **Security** → select **SQL Server and Windows Authentication mode**, then restart the SQL Server service.

4. Open `DevPulseSSE.accdb` in Access.
5. The connection setup dialog appears on first run. Enter your server address (e.g. `.\SQLEXPRESS` for a local Express instance), the login name and password you created in step 3, then click **Connect**.
6. DevPulse opens to the dashboard, pre-loaded with the sample build history.

> **Starting fresh?** Use the **Admin → Clear Data** button on the ribbon when you are ready to remove the sample data and track your own projects.

---

### Access BE Edition

No SQL Server required.

1. Download `DevPulseACE.accdb` and `DevPulse_BE.accdb` from the [Releases](../../releases) page.
2. Keep both files in the same folder.
3. Open `DevPulseACE.accdb`. DevPulse connects automatically and opens to the dashboard.

> **Starting fresh?** Use the **Admin → Clear Data** button on the ribbon when you are ready to remove the sample data and track your own projects.

## How It Was Built

See [CLAUDE.md](CLAUDE.md) for the full story — the development methodology, the tools used, and the shift from writing code to managing a project.

## License

MIT — see [LICENSE](LICENSE).
