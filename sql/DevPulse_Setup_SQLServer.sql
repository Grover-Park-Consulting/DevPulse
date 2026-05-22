-- ============================================================
-- DevPulse — SQL Server Setup Script
-- Version: 1.0  |  2026-05-19
--
-- Creates all tables, views, triggers, and seeds lookup data
-- for a fresh DevPulse installation on SQL Server.
--
-- Requirements:
--   SQL Server 2016+ or SQL Server Express (free)
--   https://www.microsoft.com/en-us/sql-server/sql-server-downloads
--
-- Instructions:
--   1. Create an empty database and switch to it:
--        CREATE DATABASE DevPulse;
--        USE DevPulse;
--   2. Run this script.
--   3. Connect the Access front end via ODBC using a SQL Server
--      login that has db_datareader + db_datawriter on this database.
--      The most recent version of the ODBC Driver for SQL Server is
--      recommended: https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server
--      See README.md for ODBC setup details.
--
-- This script is safe to run on an empty database only.
-- Running it a second time will fail on duplicate object errors.
-- ============================================================


-- ============================================================
-- LOOKUP TABLES
-- No foreign key dependencies — create first.
--
-- Naming note: the tlkp prefix ("table lookup") is a GPC
-- convention and can be changed to suit your preference
-- (e.g. lkp, ref, lu). If you rename these tables, update
-- the corresponding references in the main tables' FK
-- constraints, the two views, and the Access front end's
-- linked table names and combo box RowSources.
-- There is no single "right" naming convention — this one
-- is simply the choice made here.
-- ============================================================

CREATE TABLE tlkpProjectArea (
    ProjectAreaID   INT           IDENTITY(1,1)  NOT NULL,
    ProjectAreaName NVARCHAR(100)                NOT NULL,
    SortOrder       SMALLINT                     NOT NULL  DEFAULT 0,
    InactiveDate    DATE                         NULL,
    AccessTS        TIMESTAMP                    NOT NULL,
    CONSTRAINT PK_tlkpProjectArea PRIMARY KEY (ProjectAreaID)
);

CREATE TABLE tlkpStatus (
    StatusID     INT           IDENTITY(1,1)  NOT NULL,
    StatusName   NVARCHAR(100)                NOT NULL,
    SortOrder    SMALLINT                     NOT NULL  DEFAULT 0,
    InactiveDate DATE                         NULL,
    AccessTS     TIMESTAMP                    NOT NULL,
    CONSTRAINT PK_tlkpStatus PRIMARY KEY (StatusID)
);

CREATE TABLE tlkpStepType (
    StepTypeID   INT           IDENTITY(1,1)  NOT NULL,
    StepTypeName NVARCHAR(100)                NOT NULL,
    SortOrder    SMALLINT                     NOT NULL  DEFAULT 0,
    InactiveDate DATE                         NULL,
    AccessTS     TIMESTAMP                    NOT NULL,
    CONSTRAINT PK_tlkpStepType PRIMARY KEY (StepTypeID)
);

CREATE TABLE tlkpTool (
    ToolTypeID   INT           IDENTITY(1,1)  NOT NULL,
    ToolTypeName NVARCHAR(100)                NOT NULL,
    SortOrder    SMALLINT                     NOT NULL  DEFAULT 0,
    InactiveDate DATE                         NULL,
    AccessTS     TIMESTAMP                    NOT NULL,
    CONSTRAINT PK_tlkpTool PRIMARY KEY (ToolTypeID)
);

CREATE TABLE tlkpOutputType (
    OutputTypeID   INT           IDENTITY(1,1)  NOT NULL,
    OutputTypeName NVARCHAR(100)                NOT NULL,
    SortOrder      SMALLINT                     NOT NULL  DEFAULT 0,
    InactiveDate   DATE                         NULL,
    AccessTS       TIMESTAMP                    NOT NULL,
    CONSTRAINT PK_tlkpOutputType PRIMARY KEY (OutputTypeID)
);


-- ============================================================
-- CONFIGURATION TABLE
-- ============================================================

CREATE TABLE tblAppConfig (
    ConfigName        NVARCHAR(100)  NOT NULL,
    ConfigValue       NVARCHAR(500)  NOT NULL,
    ConfigDescription NVARCHAR(500)  NULL,
    AccessTS          TIMESTAMP      NOT NULL,
    CONSTRAINT PK_tblAppConfig PRIMARY KEY (ConfigName)
);


-- ============================================================
-- MAIN TABLES
-- Created in FK dependency order.
-- ============================================================

CREATE TABLE tblProject (
    ProjectID          INT            IDENTITY(1,1)  NOT NULL,
    ProjectName        NVARCHAR(200)                 NOT NULL,
    ProjectDescription NVARCHAR(MAX)                 NULL,
    ProjectAreaID      INT                           NULL,
    StatusID           INT                           NULL,
    StartDate          DATE                          NULL,
    CompletionDate     DATE                          NULL,
    Notes              NVARCHAR(MAX)                 NULL,
    CreatedDate        DATETIME                      NOT NULL  DEFAULT GETDATE(),
    CreatedBy          NVARCHAR(100)                 NOT NULL  DEFAULT SUSER_SNAME(),
    ModifiedDate       DATETIME                      NOT NULL  DEFAULT GETDATE(),
    ModifiedBy         NVARCHAR(100)                 NOT NULL  DEFAULT SUSER_SNAME(),
    AccessTS           TIMESTAMP                     NOT NULL,
    CONSTRAINT PK_tblProject      PRIMARY KEY (ProjectID),
    CONSTRAINT FK_tblProject_Area FOREIGN KEY (ProjectAreaID) REFERENCES tlkpProjectArea (ProjectAreaID),
    CONSTRAINT FK_tblProject_Status FOREIGN KEY (StatusID)    REFERENCES tlkpStatus (StatusID)
);

CREATE TABLE tblStep (
    StepID            INT            IDENTITY(1,1)  NOT NULL,
    ProjectID         INT                           NOT NULL,
    StepName          NVARCHAR(200)                 NOT NULL,
    StepDescription   NVARCHAR(MAX)                 NULL,
    StepTypeID        INT                           NULL,
    StatusID          INT                           NULL,
    PredecessorStepID INT                           NULL,
    StartDate         DATE                          NULL,
    CompletionDate    DATE                          NULL,
    Notes             NVARCHAR(MAX)                 NULL,
    CreatedDate       DATETIME                      NOT NULL  DEFAULT GETDATE(),
    CreatedBy         NVARCHAR(100)                 NOT NULL  DEFAULT SUSER_SNAME(),
    ModifiedDate      DATETIME                      NOT NULL  DEFAULT GETDATE(),
    ModifiedBy        NVARCHAR(100)                 NOT NULL  DEFAULT SUSER_SNAME(),
    AccessTS          TIMESTAMP                     NOT NULL,
    CONSTRAINT PK_tblStep              PRIMARY KEY (StepID),
    CONSTRAINT FK_tblStep_Project      FOREIGN KEY (ProjectID)        REFERENCES tblProject (ProjectID),
    CONSTRAINT FK_tblStep_StepType     FOREIGN KEY (StepTypeID)       REFERENCES tlkpStepType (StepTypeID),
    CONSTRAINT FK_tblStep_Status       FOREIGN KEY (StatusID)         REFERENCES tlkpStatus (StatusID),
    CONSTRAINT FK_tblStep_Predecessor  FOREIGN KEY (PredecessorStepID) REFERENCES tblStep (StepID)
);

CREATE TABLE tblSession (
    SessionID        INT            IDENTITY(1,1)  NOT NULL,
    StepID           INT                           NOT NULL,
    SessionDate      DATE                          NULL,
    SessionStartTime TIME                          NULL,
    SessionEndTime   TIME                          NULL,
    ContextReset     BIT                           NOT NULL  DEFAULT 0,
    RevisionCycles   INT                           NOT NULL  DEFAULT 0,
    Notes            NVARCHAR(MAX)                 NULL,
    CreatedDate      DATETIME                      NOT NULL  DEFAULT GETDATE(),
    CreatedBy        NVARCHAR(100)                 NOT NULL  DEFAULT SUSER_SNAME(),
    ModifiedDate     DATETIME                      NOT NULL  DEFAULT GETDATE(),
    ModifiedBy       NVARCHAR(100)                 NOT NULL  DEFAULT SUSER_SNAME(),
    AccessTS         TIMESTAMP                     NOT NULL,
    CONSTRAINT PK_tblSession       PRIMARY KEY (SessionID),
    CONSTRAINT FK_tblSession_Step  FOREIGN KEY (StepID) REFERENCES tblStep (StepID)
);

CREATE TABLE tblSessionTool (
    SessionToolID INT       IDENTITY(1,1)  NOT NULL,
    SessionID     INT                      NOT NULL,
    ToolTypeID    INT                      NOT NULL,
    UseCount      INT                      NOT NULL  DEFAULT 1,
    AccessTS      TIMESTAMP                NOT NULL,
    CONSTRAINT PK_tblSessionTool         PRIMARY KEY (SessionToolID),
    CONSTRAINT FK_tblSessionTool_Session FOREIGN KEY (SessionID)   REFERENCES tblSession (SessionID),
    CONSTRAINT FK_tblSessionTool_Tool    FOREIGN KEY (ToolTypeID)  REFERENCES tlkpTool (ToolTypeID)
);

CREATE TABLE tblOutput (
    OutputID          INT            IDENTITY(1,1)  NOT NULL,
    StepID            INT                           NOT NULL,
    OutputTypeID      INT                           NOT NULL,
    OutputCount       INT                           NOT NULL  DEFAULT 1,
    OutputDescription NVARCHAR(500)                 NULL,
    AccessTS          TIMESTAMP                     NOT NULL,
    CONSTRAINT PK_tblOutput            PRIMARY KEY (OutputID),
    CONSTRAINT FK_tblOutput_Step       FOREIGN KEY (StepID)       REFERENCES tblStep (StepID),
    CONSTRAINT FK_tblOutput_OutputType FOREIGN KEY (OutputTypeID) REFERENCES tlkpOutputType (OutputTypeID)
);


-- ============================================================
-- AUDIT TRIGGERS
-- Automatically stamp ModifiedDate and ModifiedBy on UPDATE.
-- ============================================================

GO
CREATE TRIGGER trg_tblProject_AfterUpdate
ON tblProject AFTER UPDATE AS
BEGIN
    SET NOCOUNT ON;
    UPDATE tblProject SET ModifiedDate = GETDATE(), ModifiedBy = SUSER_SNAME()
    FROM tblProject p INNER JOIN inserted i ON p.ProjectID = i.ProjectID;
END;

GO
CREATE TRIGGER trg_tblStep_AfterUpdate
ON tblStep AFTER UPDATE AS
BEGIN
    SET NOCOUNT ON;
    UPDATE tblStep SET ModifiedDate = GETDATE(), ModifiedBy = SUSER_SNAME()
    FROM tblStep s INNER JOIN inserted i ON s.StepID = i.StepID;
END;

GO
CREATE TRIGGER trg_tblSession_AfterUpdate
ON tblSession AFTER UPDATE AS
BEGIN
    SET NOCOUNT ON;
    UPDATE tblSession SET ModifiedDate = GETDATE(), ModifiedBy = SUSER_SNAME()
    FROM tblSession s INNER JOIN inserted i ON s.SessionID = i.SessionID;
END;
GO


-- ============================================================
-- VIEWS
-- ============================================================

-- vw_RptProjectSummary
-- Powers rptProjectSummary. CTE joins all project/step/session
-- data; outer query aggregates session metrics per step.
CREATE VIEW vw_RptProjectSummary AS
WITH base AS (
    SELECT
        CASE WHEN p.StatusID IN (1,2,3) THEN 0 ELSE 1 END         AS StatusGroup,
        CASE WHEN p.StatusID IN (1,2,3)
             THEN 'Open Projects' ELSE 'Closed Projects' END        AS StatusGroupLabel,
        p.ProjectID, p.ProjectName, p.ProjectDescription,
        p.Notes                                                      AS ProjectNotes,
        pa.ProjectAreaName,
        ps.StatusName                                                AS ProjectStatus,
        p.StartDate                                                  AS ProjectStartDate,
        p.CompletionDate                                             AS ProjectCompletionDate,
        s.StepID, s.StepName,
        st.StepTypeName,
        ss.StatusName                                                AS StepStatus,
        s.StartDate                                                  AS StepStartDate,
        s.CompletionDate                                             AS StepCompletionDate,
        ses.SessionID, ses.RevisionCycles, ses.ContextReset
    FROM       tblProject p
    LEFT JOIN  tblStep          s   ON p.ProjectID     = s.ProjectID
    LEFT JOIN  tlkpProjectArea  pa  ON p.ProjectAreaID = pa.ProjectAreaID
    LEFT JOIN  tlkpStatus       ps  ON p.StatusID      = ps.StatusID
    LEFT JOIN  tlkpStepType     st  ON s.StepTypeID    = st.StepTypeID
    LEFT JOIN  tlkpStatus       ss  ON s.StatusID      = ss.StatusID
    LEFT JOIN  tblSession       ses ON s.StepID        = ses.StepID
)
SELECT
    StatusGroup, StatusGroupLabel,
    ProjectID, ProjectName, ProjectDescription, ProjectNotes,
    ProjectAreaName, ProjectStatus, ProjectStartDate, ProjectCompletionDate,
    StepID, StepName, StepTypeName, StepStatus, StepStartDate, StepCompletionDate,
    COUNT(SessionID)                                                 AS SessionCount,
    ISNULL(SUM(RevisionCycles), 0)                                  AS RevisionTotal,
    ISNULL(SUM(CASE WHEN ContextReset = 1 THEN 1 ELSE 0 END), 0)   AS ResetCount
FROM base
GROUP BY
    StatusGroup, StatusGroupLabel,
    ProjectID, ProjectName, ProjectDescription, ProjectNotes,
    ProjectAreaName, ProjectStatus, ProjectStartDate, ProjectCompletionDate,
    StepID, StepName, StepTypeName, StepStatus, StepStartDate, StepCompletionDate;
GO

-- vw_ChartStepsLong
-- Long-format view for the Progress by Project chart on frmMain.
-- Returns one row per project per category (Completed / Open).
CREATE VIEW vw_ChartStepsLong AS
SELECT p.ProjectName,
       'Completed Steps'                                    AS StepCategory,
       COUNT(CASE WHEN s.StatusID = 4 THEN 1 END)          AS StepCount
FROM tblProject p
INNER JOIN tblStep s ON p.ProjectID = s.ProjectID
WHERE p.StatusID IN (1, 2, 3)
GROUP BY p.ProjectName
UNION ALL
SELECT p.ProjectName,
       'Open Steps'                                         AS StepCategory,
       COUNT(CASE WHEN s.StatusID <> 4 THEN 1 END)         AS StepCount
FROM tblProject p
INNER JOIN tblStep s ON p.ProjectID = s.ProjectID
WHERE p.StatusID IN (1, 2, 3)
GROUP BY p.ProjectName;
GO


-- ============================================================
-- SEED DATA — LOOKUP TABLES
-- Insert in SortOrder sequence; IDENTITY assigns IDs.
-- Active rows only — soft-deleted test entries excluded.
-- ============================================================

-- Project Areas
INSERT INTO tlkpProjectArea (ProjectAreaName, SortOrder) VALUES
    ('Access',        10),
    ('PowerApps',     20),
    ('SQL Server',    30),
    ('Web',           40),
    ('Cross-Domain',  50),
    ('Other',         99);

-- Statuses
INSERT INTO tlkpStatus (StatusName, SortOrder) VALUES
    ('Planning',    10),
    ('In Progress', 20),
    ('On Hold',     30),
    ('Complete',    40),
    ('Cancelled',   50);

-- Step Types
INSERT INTO tlkpStepType (StepTypeName, SortOrder) VALUES
    ('Planning',       10),
    ('Development',    20),
    ('Testing',        30),
    ('Documentation',  40),
    ('Research',       50),
    ('Refactoring',    60);

-- Tools (Claude Code listed first — primary tool for AI-assisted development)
INSERT INTO tlkpTool (ToolTypeName, SortOrder) VALUES
    ('Claude Code',        10),
    ('Access Explorer MCP', 20),
    ('SQL MCP',            30),
    ('Claude Chat',        40),
    ('PowerShell',         50),
    ('VSCode',             60),
    ('Access',             70),
    ('Other',              99);

-- Output Types
INSERT INTO tlkpOutputType (OutputTypeName, SortOrder) VALUES
    ('Schema Design',  10),
    ('SQL Script',     20),
    ('VBA Module',     30),
    ('Form',           40),
    ('Report',         50),
    ('Wiki Entry',     60),
    ('Documentation',  70),
    ('Other',          99);


-- ============================================================
-- SEED DATA — APP CONFIGURATION
-- ============================================================

INSERT INTO tblAppConfig (ConfigName, ConfigValue, ConfigDescription) VALUES
    ('AppTitle',           'DevPulse', 'Application display title'),
    ('AppVersion',         '1.0',      'Application version'),
    ('AvgSessionMinutes',  '45',       'Default average session length in minutes for estimates');



-- ============================================================
-- DevPulse Sample Data: Core Build (Tasks 1-21)
-- ============================================================

SET IDENTITY_INSERT tblProject ON;
INSERT INTO tblProject (ProjectID,ProjectName,ProjectDescription,ProjectAreaID,StatusID,StartDate,CompletionDate,Notes) VALUES (1,'DevPulse','AI-assisted development performance tracker — measures and publishes session metrics for blog, YouTube, and Access user group',5,2,'2026-05-08',NULL,NULL);
SET IDENTITY_INSERT tblProject OFF;

SET IDENTITY_INSERT tblStep ON;
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1,1,'Initial Design — Goals, Requirements, Schema Approval','Design session establishing project goals, requirements, and SQL Server schema',1,4,NULL,'2026-05-08','2026-05-08',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (2,1,'SQL Server Backend Build — Database, Tables, Triggers, SPs, Seeds','Created DevPulse database on gpconsultingbb; 11 tables, 3 AFTER UPDATE triggers, 5 usp_AddUpdate SPs, seeded all lookup tables and tblAppConfig',2,4,1,'2026-05-09','2026-05-09',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (3,1,'Create Access Frontend — accdb and CodeArchive Reference','Created DevPulse.accdb at D:\GroverPark\Devpulse; added CodeArchive-GPG201964B_Prod reference',2,4,2,'2026-05-09','2026-05-09',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (4,1,'Access Infrastructure — USysConnect, USysAppString, USysRibbons, tblLookupCacheConfig, tmpl_ tables','Created 9 local Access tables: USysConnect, USysAppString, USysRibbons, tblLookupCacheConfig, 5 tmpl_ lookup templates. Seeded connection, app strings, and cache config.',2,4,3,'2026-05-09','2026-05-09',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (5,1,'Access Infrastructure — lcl_ Lookup Cache Tables','Created 5 lcl_ local lookup cache tables: lcl_tlkpProjectArea, lcl_tlkpStatus, lcl_tlkpStepType, lcl_tlkpTool, lcl_tlkpOutputType',2,4,4,'2026-05-09','2026-05-09',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (6,1,'modAppSetup — InitApp, LinkSQLTables, BuildConnStr','Created modAppSetup with InitApp (links 11 SQL Server tables via ODBC Driver 18, calls codearchive.InitApplication and codearchive.CreateLookupCache). Driver 18 with TrustServerCertificate=yes for local server.',2,4,5,'2026-05-09','2026-05-09',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (7,1,'InitApp() Testing, Bug Fixes, Git Setup','Three fixes during InitApp() testing: (1) AppLogo empty-string guard in CodeArchive, (2) dbAttachSavePWD on linked tables, (3) connection string reorder with Authentication=SqlPassword and tcp: prefix. Git repo initialized, first commit made.',3,4,6,'2026-05-09','2026-05-09',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (9,1,'frmProject + sfrmStep — Build and Test','Built frmProject with sfrmStep continuous subform; project search combo, navigation buttons, and start/completion date fields wired up.',2,4,7,'2026-05-10','2026-05-10','Search combo with UNION ALL/USysAppString pattern. TempVar filter. Conditional nav buttons. Committed 5439521.');
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (10,1,'frmStep + sfrmSession + sfrmOutput + sfrmTool — Build','Built frmStep with four linked subforms: sfrmSession, sfrmOutput, sfrmTool, and sfrmSessionTool for tracking step details.',2,4,9,'2026-05-10','2026-05-10','Button-switch lazy loading via single sfrmActive subform. SourceObject swapped on click. Three subforms. Committed 1fae597.');
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (11,1,'Bug Fixes — Ghost Text, VBA Array Pattern, RecordSelectors','Fixed ghost text on combo boxes, corrected VBA array pattern for the linked table list, and suppressed RecordSelectors on all subforms.',6,4,9,'2026-05-10','2026-05-10','ColumnWidths=0;2880 on all search combos. RecordSelectors removed from UpdateNavButtons (form-wide reflow causes Header rendering artifact). HighlightActiveButton rewritten with typed CommandButton array. Bug reported to MS Access team email list. Committed 714cf43.');
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1002,1,'frmSession — Design and Planning','Reconsidered frmSession purpose and design. Key decisions: continuous view with footer tool subform; LinkMasterFields/LinkChildFields confirmed valid in form footer; TempVar-based CF for current-row highlight; single-row header; planning and build split into separate steps.',1,4,11,'2026-05-15','2026-05-15',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1003,1,'frmSession — Build','Build qrySession, sfrmSessionTool, and frmSession per approved design: cascading header combos, continuous detail view, footer tool subform via LinkMasterFields/LinkChildFields, TempVar conditional formatting for current-row highlight.',2,4,1002,'2026-05-15','2026-05-15',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1004,1,'Lookup Management Forms + frmAppConfig — Build and Test','Built 5 lookup management forms (frmLookupTool, frmLookupStatus, frmLookupOutputType, frmLookupProjectArea, frmLookupStepType) and frmAppConfig. Fixed soft-delete cache refresh bug: added CloseLookupForm() to modAppSetup, which truncates the lcl_tlkp* table before calling CreateLookupCache to force a count mismatch and guarantee refresh.',2,4,1003,'2026-05-15','2026-05-15','');
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1005,1,'frmMain Dashboard — Scaffold and Test','Built frmMain as ribbon-driven dashboard (no navigation buttons). Open Projects and Completed Projects subforms using qryOpenProjects/qryCompletedProjects against linked SQL Server tables. Five live metrics in ''What We''ve Done'' panel. Fixed two build issues: GROUP BY + COUNT against linked tables silently drops rows in Jet/ACE — replaced with DCount expression; section heights defaulted to 5 inches — fixed via temporary VBA helper.',2,4,1004,'2026-05-15','2026-05-15','');
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1006,1,'rptProjectSummary — Design and Planning','Defined report structure: two-level grouping (Open/Closed then by Project), project header with name/area/status/dates/description/notes, step detail rows, project footer with Steps/Sessions/Revision Cycles/Context Resets totals using DCount/DSum per StepID summed in footer to avoid linked-table GROUP BY issue.',1,4,1005,'2026-05-15','2026-05-15','');
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1007,1,'rptProjectSummary — Build and Test','Built rptProjectSummary via PowerShell UTF-16 LE + access_import_text (LoadFromText). Discovered correct Access report text schema from WorkEffort rptInvoiceHourly: Version=21, Begin BreakLevel for groups, Begin BreakHeader/BreakFooter for sections. Nine iterations to reach approved layout: Calibri font, navy title, landscape 14400 width, CanGrow/CanShrink on group header, right-aligned column date headers, bold values / normal labels.',2,4,1006,'2026-05-16','2026-05-16','');
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1008,1,'rptProjectSummary — SQL Server View & Performance Optimization','Created vw_RptProjectSummary CTE view on SQL Server; rewrote report to use pre-aggregated view, reducing render time significantly.',6,4,NULL,'2026-05-16','2026-05-16',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1009,1,'frmMain Dashboard — Chart (Task 15)','Built chrtStepsByProject horizontal stacked bar chart on frmMain using vw_ChartStepsLong long-format SQL Server view.',2,4,NULL,'2026-05-16','2026-05-16',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1011,1,'Task 16 — AutoExec, startup relink fix, Ribbon, AutoCenter','Built AutoExec macro, USysLinkedTables-driven relink routine, Ribbon XML with four groups, modRibbon navigation dispatcher, and BackupDevPulse function.',2,4,NULL,'2026-05-17','2026-05-18',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1012,1,'Task 17 — Control highlighting (FormControlHighlighting + CF)','Wired FormControlHighlighting class on 9 single-record forms and injected conditional formatting on 7 continuous subforms for visual focus feedback.',2,4,NULL,'2026-05-18','2026-05-18',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1013,1,'Task 17 (amendment) — Correct continuous form highlighting','Corrected 6 lookup forms from single-record to continuous-view highlighting pattern; established rule: single-record uses class, continuous uses conditional formatting.',6,4,NULL,'2026-05-18','2026-05-18',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1014,1,'Task 18 — Error handlers on all form procs','Added full errHandler blocks using VBE reflection pattern to all substantive form procedures; fixed ApplyCF highlight color in frmSession.',2,4,NULL,'2026-05-18','2026-05-18',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1015,1,'Task 19 — Distribution Prep: AppTitle, AppIcon, startup options','Set AppTitle, AppIcon, StartupForm, and startup options; left developer access tools enabled — audience is Access developers exploring the app.',4,4,NULL,'2026-05-18','2026-05-18',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1018,1,'Testing — end-to-end UI test and bug fixes','End-to-end UI test found 6 bugs: missing AccessTS timestamps on all SQL Server tables, date field alignment, subform sizing, combo requery, soft-delete filter, and report spacing.',3,4,NULL,'2026-05-18',NULL,NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1019,1,'Post-v1 layout fixes - subform margins, control widths, lookup form columns','Systemic layout audit across all subforms and lookup forms; established subform right-margin rule and container-width formula; narrowed date and text columns throughout.',6,4,NULL,'2026-05-19','2026-05-19',NULL);
SET IDENTITY_INSERT tblStep OFF;

SET IDENTITY_INSERT tblSession ON;
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1,1,'2026-05-08','13:38:00','14:15:00',0,0,'Design session establishing project goals, data model (Project-Step-Session hierarchy), and toolchain; approved schema and development approach.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (2,2,'2026-05-09','08:30:00',NULL,0,1,'Created DevPulse SQL Server database with 11 tables, 3 audit triggers, and 5 stored procedures; 1 revision cycle to correct trigger syntax.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (3,3,'2026-05-09','08:30:00',NULL,0,0,'Created DevPulse.accdb, added CodeArchive external library reference, and confirmed VBA compilation with no errors.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (4,4,'2026-05-09','08:30:00',NULL,0,0,'Built 9 local Access tables including USysConnect, USysAppString, USysRibbons, tblLookupCacheConfig, and 5 tmpl_ seed tables.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (5,5,'2026-05-09','08:30:00',NULL,0,0,'Created 5 lcl_ lookup cache tables; confirmed MaintainLookupCache populates all caches correctly on first call.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (6,6,'2026-05-09','08:30:00',NULL,0,3,'Built modAppSetup with BuildConnStr, LinkSQLTables (functional relink test), and InitApp; 3 revision cycles to resolve ODBC connection string format.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (7,7,'2026-05-09','08:30:00','11:40:00',0,3,'Fixed 3 InitApp bugs: AppLogo empty-string crash, VBA array initialization error, and linked table ordering issue.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1002,9,'2026-05-10',NULL,NULL,0,0,'Retrospective entry — session times not recorded.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1003,10,'2026-05-10',NULL,NULL,0,0,'Retrospective entry — session times not recorded.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1004,11,'2026-05-10',NULL,NULL,0,0,'Retrospective entry — session times not recorded.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1005,1002,'2026-05-15','13:25:00','15:35:00',1,5,'Context reset mid-session (accidental window close). Design rethought from initial concept: purpose clarified, LinkMasterFields footer claim corrected, CF approach debated and resolved, header collapsed to single row.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1006,1003,'2026-05-15','15:47:00','18:07:00',0,5,'Built qrySession, sfrmSessionTool, frmSession. CF persistence solved via acDesign mode; Me.Refresh (not Me.Repaint) required for CF re-evaluation in continuous view (fix by GH).');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1007,1004,'2026-05-15','09:00:00','09:45:00',1,1,'Built 5 lookup management forms and frmAppConfig. Context reset (VS Code crash) lost first session; work reconstructed from memory. Bug found in testing: soft delete did not refresh lcl_tlkp* cache — CreateLookupCache skips refresh when counts are equal and InactiveDate change leaves count unchanged. Fix: CloseLookupForm(sLocalTable) added to modAppSetup; truncates local table before cache rebuild to guarantee count mismatch. All 5 Form_Close procedures simplified to single Call.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1008,1005,'2026-05-15','10:10:00','11:53:00',0,2,'Built frmMain scaffold: sfrmOpenProjects, sfrmCompletedProjects, qryOpenProjects, qryCompletedProjects, frmMain. Two revision cycles: (1) GROUP BY + COUNT against linked SQL Server tables silently dropped rows in Jet/ACE query engine — replaced StepCount with DCount expression; (2) section heights defaulted to 5 inches — fixed via temporary FixSectionHeights() VBA helper. Metrics panel label changed from ''Metrics'' to ''What We''ve Done''.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1009,1006,'2026-05-15','12:01:00','12:10:00',0,0,'Defined rptProjectSummary structure: two-level grouping (Open/Closed outer, ProjectID inner), project header fields, step detail columns, footer totals pattern using hidden DCount/DSum per step summed in group footer.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1010,1007,'2026-05-16','09:07:00','12:15:00',1,9,'Built rptProjectSummary. Key discovery: Access report text format uses BreakLevel/BreakHeader/BreakFooter (not GroupLevel/GroupHeader). WorkEffort rptInvoiceHourly was the Rosetta Stone for the schema. Nine layout iterations: font (Calibri), title color, CanGrow/CanShrink, landscape width (14400), column header right-alignment, normal labels/bold values. Context reset (VS Code crash). One minor manual tweak by GH (100% rule exception).');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1011,1008,'2026-05-16',NULL,NULL,1,1,'Replaced direct table joins with CTE view; 1 revision cycle to correct CTE column aliases causing report grouping errors.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1012,1009,'2026-05-16',NULL,NULL,1,7,'Built chart using long-format view and ChartLegend/PIVOT pattern; 7 revision cycles to resolve Access ModernChart legend prefix issue.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1014,1011,'2026-05-18',NULL,NULL,1,11,'Highest revision cycle count in the build (11); root complexity was Ribbon XML syntax and the relink functional-test pattern. Context reset required mid-session.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1015,1012,'2026-05-18',NULL,NULL,0,8,'Wired FormControlHighlighting class and conditional formatting; 8 cycles to resolve correct form type classification (single-record vs continuous).');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1016,1013,'2026-05-18',NULL,NULL,0,3,'Corrected form type misclassification on 6 lookup forms; established authoritative rule distinguishing single-record from continuous highlighting approach.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1017,1014,'2026-05-18',NULL,NULL,0,2,'Added errHandler blocks to 12 form procedures using VBE reflection pattern; eliminates hardcoded module and procedure name strings.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1018,1015,'2026-05-18',NULL,NULL,0,1,'Set startup properties; deliberately left AllowSpecialKeys and AllowFullMenus ON — DevPulse is for Access developers who need to explore its internals.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1020,1018,'2026-05-18',NULL,NULL,0,4,'Team lead must evaluate the application and all bug fixes and data entry');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1021,1019,'2026-05-19',NULL,NULL,0,6,'Narrowed columns across 7 subforms and 5 lookup forms; 6 revision cycles to resolve shared-container constraints and right-margin calculations.');
SET IDENTITY_INSERT tblSession OFF;

SET IDENTITY_INSERT tblSessionTool ON;
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1,1,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (2,3,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (3,4,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (4,5,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (5,6,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (6,7,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1002,1020,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1003,1,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1004,2,2,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1005,2,4,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1006,1006,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1007,1006,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1008,1006,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1009,1004,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1010,1004,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1011,1004,4,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1012,2,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1013,1003,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1014,1002,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1015,1005,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1016,1007,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1017,1008,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1018,1009,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1019,1010,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1020,1011,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1021,1012,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1022,1014,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1023,1015,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1024,1016,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1025,1017,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1026,1018,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1027,1020,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1028,1021,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1029,1021,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1030,1020,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1031,1018,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1032,1017,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1033,1016,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1034,1015,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1035,1014,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1036,1012,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1037,1011,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1038,1010,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1039,1009,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1040,1008,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1041,1007,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1042,1005,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1043,1002,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1044,1003,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1045,6,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1046,7,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1047,2,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1048,3,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1049,4,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1050,5,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1051,5,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1052,4,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1053,3,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1054,2,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1055,7,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1056,6,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1057,1003,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1058,1004,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1059,1002,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1060,1005,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1061,1007,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1062,1008,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1063,1009,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1064,1010,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1065,1011,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1066,1012,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1067,1014,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1068,1015,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1069,1016,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1070,1017,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1071,1018,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1072,1021,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1073,1,1003,1);
SET IDENTITY_INSERT tblSessionTool OFF;


-- ============================================================
-- Setup complete.
-- Next step: connect the Access front end (DevPulseSSE.accdb)
-- via ODBC using frmSetupConnection, or open DevPulseACE.accdb
-- for the Access backend edition.
-- See README.md for full setup instructions.