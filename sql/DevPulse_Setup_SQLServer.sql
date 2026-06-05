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
    SessionStartTime DATETIME                      NULL,
    SessionEndTime   DATETIME                      NULL,
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
-- Post-v1 additional records (Project 1 — DevPulse app)
-- Tasks added after v1.0.0 release: 2026-05-24 through 2026-05-27
-- ============================================================
SET IDENTITY_INSERT tblStep ON;
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1041,1,'Post-v1 Query fixes','Fixed logic error in frmSession recordsource query surfaced during post-v1 testing; applied update to both distro editions in the repo.',6,4,NULL,'2026-05-24','2026-05-24',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1043,1,'frmSession — Testing Bug','Discovered and corrected design flaw in two DateTime fields; required changes to frmSession form and the underlying table schema.',6,4,NULL,'2026-05-25','2026-05-25',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1050,1,'Released Video and posted Forum announcements','Drafted and posted platform-specific DevPulse launch announcements across YouTube, LinkedIn, X/Twitter, Reddit, and Access community forums.',4,4,NULL,'2026-05-27','2026-05-27',NULL);
SET IDENTITY_INSERT tblStep OFF;

SET IDENTITY_INSERT tblSession ON;
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1038,1041,'2026-05-24','15:30:00','16:00:00',0,0,'Testing. Surfaced logic error in frmSession''s recordsource query. Gave Claude the fix to update the Distro versions in the repo.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1044,1043,'2026-05-25','05:00:00','06:00:00',0,0,'Testing. Surfaced a design flaw in two DateTime fields that required changes in frmSession as well as the table in the database.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1047,1050,'2026-05-27','07:30:00','09:00:00',0,0,'Requested announcements tailored to different venues: X, YouTube, LinkedIn, Forums, Reddit.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1048,1050,'2026-05-27','09:00:00','10:30:00',0,0,'Posted and edited announcements.');
SET IDENTITY_INSERT tblSession OFF;

SET IDENTITY_INSERT tblSessionTool ON;
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1078,1044,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1079,1044,1,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1080,1044,1003,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1081,1044,4,1);
SET IDENTITY_INSERT tblSessionTool OFF;


-- ============================================================
-- Launch Materials (ProjectID=1002) + Forum Announcements (ProjectID=1003)
-- Steps and sessions added: 2026-05-18 through 2026-06-05
-- Includes all DevPulse launch-phase work: distro build, GitHub release,
-- blog posts, YouTube videos (demo, PM lessons, tool stack walkthrough),
-- and Access Pacific User Group presentation.
-- Data corrections applied: StepID=1051 StepTypeID fixed 1003→4;
-- ProjectID=1003 StatusID fixed 1→4.
-- ============================================================
SET IDENTITY_INSERT tblProject ON;
INSERT INTO tblProject (ProjectID,ProjectName,ProjectDescription,ProjectAreaID,StatusID,StartDate,CompletionDate,Notes) VALUES (1002,'Prepare launch Materials for sharing DevPulse','Work with Claude to create a distro via GitHub. Package with script for SQL Server db and also with an accdb be.',5,2,'2026-05-18',NULL,NULL);
INSERT INTO tblProject (ProjectID,ProjectName,ProjectDescription,ProjectAreaID,StatusID,StartDate,CompletionDate,Notes) VALUES (1003,'Notify various user forums of availability','Press release time notices for AWF, and others',5,4,'2026-05-18',NULL,NULL);
SET IDENTITY_INSERT tblProject OFF;

SET IDENTITY_INSERT tblStep ON;
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1016,1002,'Create a PRD',NULL,1,4,NULL,'2026-05-18','2026-05-19',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1017,1003,'Prepare list of user forums',NULL,1,4,NULL,'2026-05-18','2026-05-27',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1020,1002,'SQL Server setup script — schema, seed data, views, no credentials',NULL,2,4,NULL,'2026-05-19','2026-05-19',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1021,1002,'Access BE version — convert FE to local accdb backend',NULL,2,4,NULL,'2026-05-19','2026-05-20',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1022,1002,'Clear example data utility — script or Admin button',NULL,2,4,NULL,'2026-05-19',NULL,NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1023,1002,'GitHub — public repo, dual-backend structure, Releases for binaries',NULL,2,1,NULL,'2026-05-19',NULL,NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1024,1002,'README — story first, setup second',NULL,4,1,NULL,'2026-05-19',NULL,NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1025,1002,'Architecture document — technical stack and development methodology',NULL,4,1,NULL,'2026-05-19',NULL,NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1026,1002,'Blog post — narrative first, data as argument; mutual learning story',NULL,4,1,NULL,'2026-05-19',NULL,NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1027,1002,'YouTube — demo walkthrough',NULL,4,1,NULL,'2026-05-19',NULL,NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1029,1002,'MaintainLookupCache portability fix: stale links + DCount',NULL,2,4,NULL,'2026-05-21','2026-05-21',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1031,1002,'Strip CodeArchive from DevPulseSSE + scrub credentials from CLAUDE.md',NULL,2,4,NULL,'2026-05-21','2026-05-21',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1032,1002,'First-run connection setup for DevPulseSSE','Built frmSetupConnection modal popup and modSQLConnect (InitConnect, FixConnections, StrConnectionString, GenerateIndexSQL) to guide first-run users through SQL Server credential setup and table linking; removed StartupForm db property and fixed dbSeeChanges misuse in QueryDef.OpenRecordset.',2,4,NULL,'2026-05-21','2026-05-21',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1033,1002,'GitHub — public repo, dual-backend structure, Releases for binaries','Completed all pre-push preparation for the GitHub public repo: IP audit, data enrichment (step descriptions, session notes, tool records), accdb cleanup, SQL setup script updated with Tasks 1-21 sample data, Access BE populated, public CLAUDE.md drafted, MIT LICENSE added.',2,4,NULL,'2026-05-22','2026-05-22',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1034,1002,'GitHub initial push — assembly, modTerminate removal, access_ace rename','Pushed initial public release to Grover-Park-Consulting/DevPulse (261 files); caught and removed modTerminate.bas from SSE source post-push; renamed access_be to access_ace in both repos; created assemble.ps1 for future sync updates.',2,4,NULL,'2026-05-22','2026-05-22',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1035,1002,'README — story first, setup second','Expanded the public README from a stub to complete documentation with full setup instructions for both SQL Server and Access BE editions, including SQL Server authentication configuration and login creation. Also merged conflicting GitHub draft edits into the local distro CLAUDE.md and established the pull-then-overwrite recovery pattern for future sync conflicts.',4,4,NULL,'2026-05-23','2026-05-23',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1036,1002,'Architecture document — technical stack and development methodology','Created ARCHITECTURE.md as a technical companion to CLAUDE.md in the public repo, covering the data model (Mermaid ER diagram), backend editions, VCS rationale for AI-assisted development, and toolchain. Updated assemble.ps1 to include the new file in future distro syncs.',4,4,NULL,'2026-05-23','2026-05-23',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1037,1002,'Blog post — narrative first, data as argument; mutual learning story','Wrote and published blog post "Access and AI: What 11 Tries on One Task Taught Me" — a forensic account of the DevPulse build arguing that AI-assisted development discipline matters more than specific tools, grounded in revision cycle data with Task 16 (11 cycles, Ribbon XML) as the canonical illustration.',4,4,NULL,'2026-05-23','2026-05-23',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1038,1002,'Complete distro packaging','Applied Clear Data utility to DevPulse_Prod.accdb and rebuilt DevPulseSSE.accdb distro from scratch after distro state was lost; committed all three edition VCS exports (Prod, ACE, SSE) under correct folder names and removed obsolete DevPulse.accdb.src.',2,4,NULL,'2026-05-24','2026-05-24',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1039,1002,'GitHub Releases — binary asset upload','Published GitHub Release v1.0.0 with four downloadable assets: DevPulseSSE.accdb, DevPulseACE.accdb, DevPulse_BE.accdb, and DevPulse_Setup_SQLServer.sql; release live at github.com/Grover-Park-Consulting/DevPulse/releases/tag/v1.0.0.',2,4,NULL,'2026-05-24','2026-05-24',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1040,1002,'qrySession cross-project session filter fix + distro repo sync','Fixed qrySession to filter sessions by both StepID and ProjectID — original StepID-only filter caused cross-project session bleed in the show-all case (lngStepID=0). Joins tblStep to bring ProjectID into scope; adds lngProjectID TempVar condition as second AND predicate. Applied to Prod, ACE, and SSE editions; frmSession control widths adjusted; distro repo synced.',3,4,NULL,'2026-05-24','2026-05-24',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1042,1002,'tblSession DateTime migration + frmSession time control and SessionID fixes','Migrated tblSession.SessionStartTime and SessionEndTime from Time(7) to DateTime for Access compatibility; updated frmSession in all three editions with correct Format/InputMask and a physical txtSessionID bound control; patched Form_Current VBA to reference Me.txtSessionID; updated setup script for new installs.',3,4,NULL,'2026-05-25','2026-05-25',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1044,1002,'YouTube — demo walkthrough','Recorded DevPulse Starter YouTube demo using the Production version (includes stub data for three additional projects for a fuller Dashboard view); product walkthrough, trailer for planned toolset follow-up videos (Claude Code, JoyfulService VCS, VS Code, Access, SQL Server), and call to action for developer feedback toward DevPulse Pro; video notes that the distro version includes only DevPulse project data.',4,4,NULL,'2026-05-25','2026-05-27',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1045,1003,'Announcement posts — five platform versions','Created five platform-specific announcement posts (forum, X, LinkedIn, YouTube description, Reddit r/MSAccess) for DevPulse Starter launch; grill session drove the frame from product announcement to data-led PM-judgment argument; all versions saved to .distro/announcement-posts.md pending YouTube URL.',4,4,NULL,'2026-05-27','2026-05-27',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1051,1003,'Post announcements',NULL,4,4,NULL,'2026-05-27','2026-05-27',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1052,1003,'Post to forums','Posted platform-specific DevPulse Starter launch announcements to two Access community forums, Reddit (r/MSAccess), LinkedIn, and X/Twitter; YouTube description was posted at upload time. Active public Access forum ecosystem has dwindled significantly.',4,4,NULL,'2026-05-27','2026-05-27',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1065,1002,'YouTube -- PM lessons learned (script)','Developed full recording script for YouTube PM lessons learned video covering six sections: opening/series framing, Tom''s email, Wiki story, three PM behavior change points anchored in session data, reflections, and close. Script saved to .grill/ for on-screen use during recording.',4,4,NULL,'2026-05-29','2026-05-29',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1066,1002,'YouTube -- PM lessons learned (filming and upload)','Filmed and edited the PM lessons learned YouTube video covering six sections (opening/series framing, Tom''s email, Wiki story, three PM behavior change points, reflections, close); uploaded to YouTube at https://youtu.be/dC-TJrjU7hA',4,4,NULL,'2026-05-30','2026-05-30',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1067,1002,'Blog post — PM lessons learned (from video script)','Convert the PM lessons learned YouTube video script into a blog post for gpcdata.com; includes screenshots if supported by the CMS.',4,4,NULL,'2026-05-31','2026-05-31',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1068,1002,'YouTube — tool stack walkthrough (filming and upload)','Filmed and uploaded the tool stack walkthrough YouTube video; 9 segments, 44:50 final runtime; two deliberate context resets during MusicAlbums demo to illustrate context management on camera.',4,4,NULL,'2026-05-31','2026-06-05',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1069,1002,'Distro refresh — dual YouTube links + Project 1 seed data','Added ''Watch the Build'' section to the public GitHub README referencing both published YouTube videos, and refreshed the SQL Server setup script and DevPulse_BE.accdb with three post-v1 Project 1 steps and four sessions logged after the v1.0.0 release (StepIDs 1041, 1043, 1050).',2,4,NULL,'2026-05-31','2026-05-31',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1070,1002,'YouTube — tool stack walkthrough (script)','Script reviewed, revised, and approved for the tool stack walkthrough video. What: Approved recording script at .grill/tool-stack-video-script.md. Why: A finalized script is the prerequisite for on-camera recording — structure and narration must be settled before the camera rolls. How 1: George edited the script directly in VS Code — revised hook, demo setup, closing disclosure (tblTrack staged admission), and tool stack framing. How 2: Claude applied targeted fixes: blog link narration, spacing typos, and full 10-tool PPT slide added to closing.',4,4,NULL,'2026-06-01','2026-06-01',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (1071,1002,'YouTube — tool stack walkthrough (pre-recording cleanup)','Removed pre-build artifacts from the tool stack video run-through to restore a clean demo environment for on-camera recording. Deleted MusicAlbums.accdb and dropped the Albums schema (tblAlbum, tblAlbumGenre, tblAlbumPersonnel, tblAlbumRole, tblTrack) from DevPulse_affinity on GPGEORGEMOBILE\GPMOBILE2019. Both confirmed clean.',2,4,NULL,'2026-06-01','2026-06-01',NULL);
INSERT INTO tblStep (StepID,ProjectID,StepName,StepDescription,StepTypeID,StatusID,PredecessorStepID,StartDate,CompletionDate,Notes) VALUES (3075,1003,'Access Pacific User Group — DevPulse overview','Presented DevPulse overview to Access Pacific User Group; shared GitHub repo link and explained context resets and revision cycles.',4,4,NULL,'2026-06-04','2026-06-04',NULL);
SET IDENTITY_INSERT tblStep OFF;

SET IDENTITY_INSERT tblSession ON;
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1022,1020,'2026-05-19','15:00:00','15:35:00',0,0,'First pass — queried live schema, generated clean distro script. Removed test entries, reordered tlkpTool, reset AppVersion to 1.0.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1023,1016,'2026-05-19',NULL,NULL,0,0,'Planning completed. Decisions: both SQL Server Express and Access BE backends; SSE is primary. Example DevPulse data included in distro. 8 tasks defined (StepIDs 1020-1027). CodeArchive dependency handled via modCodeLibraryExtracts import. Content themes captured for blog, architecture doc, YouTube.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1024,1021,'2026-05-20','06:11:00','06:12:00',1,1,NULL);
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1025,1022,'2026-05-20','11:20:00','14:10:00',1,7,'Form Design. Built Clear Example Data utility: frmClearExampleData + sfrmClearAppConfig, ribbon CLEAR DATA button (Admin group, node 22), HOME button (NAVIGATE group, node 1), getEnabled callback. Multiple revision cycles on ribbon getter callback signature — correct pattern is Sub + ByRef Variant output, confirmed from NW25-Dev reference. Design standards fixes: FontSize=14, section heights, right-margin compliance.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1027,1029,'2026-05-21','12:29:00','12:52:00',0,1,NULL);
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1028,1031,'2026-05-21','13:10:00','19:40:00',0,8,'Missing. Gap in our work rules that require Notes for all steps');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1029,1032,'2026-05-21','19:53:00','20:35:00',1,2,'Key fixes: StartupForm db property deleted via DAO Properties.Delete (Application.SetOption rejects it); dbSeeChanges invalid as LockEdit in QueryDef.OpenRecordset — trimmed to (dbOpenSnapshot, dbSQLPassThrough). DevPulseSSE relinked to GPC SQL Server during testing — 13 linked tables must be deleted before GitHub deployment.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1030,1033,'2026-05-22','07:02:00','10:10:00',0,2,'IP audit found modTerminate (Philipp Steifel, private communication) as only significant third-party code — removed as dead code with no replacement needed. Access BE tlkpTool realigned to SQL Server IDs (1002=Claude Code, 1003=Access) after FK mismatch caused silent insert failures in first batch run.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1031,1034,'2026-05-22','10:26:00','11:25:00',0,2,'modTerminate.bas was in DevPulseSSE.accdb.src/modules/ but not in access_sse/ — grepping the reference copy folder missed it; lesson is to always audit the full accdb.src folder. access_be naming caught by George after seeing the live repo; renamed to access_ace to mirror access_sse convention.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1032,1035,'2026-05-23','09:36:00','10:05:00',0,2,'SQL Server edition requires mixed-mode authentication and an explicit login — documented with example T-SQL and an Express-specific troubleshooting note. Direct GitHub edits conflict with the assemble.ps1 workflow; recovery pattern is to discard staging uncommitted changes, pull from GitHub, then re-run assemble.ps1.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1033,1036,'2026-05-23','10:15:00','11:53:00',0,1,'Grill session clarified that the architecture doc and blog post serve genuinely different audiences and must be separate documents — architecture doc is a neutral technical reference, blog carries the personal narrative and community debate. Acknowledgment of colleagues who shaped the VCS argument reserved for the blog where the full discussion can be told properly.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1034,1037,'2026-05-23','12:08:00','13:05:00',1,1,'Grill session confirmed blog is distinct from architecture doc (different audience, different content); post inserted directly into production SQL Server blog database via PowerShell Invoke-Sqlcmd rather than copy/paste, a workflow that fits the blog''s own thesis about AI-assisted development.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1035,1038,'2026-05-24','09:00:00','10:30:00',1,3,'Root cause of missing SSE distro: distro modifications were applied to production in place, then frmSetupConnection reconnected to SQL Server and overwrote the distro state; fix is to copy accdb to distro folder first, then modify the copy. Surgical VBE tools (access_vbe_patch_proc + access_vbe_append) are safer than full form replacement for targeted changes to existing forms.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1036,1039,'2026-05-24','11:00:00','11:30:00',0,0,'Release notes drafted from real build data context (25 sessions, 72 revision cycles); assets confirmed present and sizes verified before upload; George uploaded via GitHub web UI after gh CLI was not found in PATH.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1037,1040,'2026-05-24','11:47:00','12:10:00',0,0,'Bug was invisible with a single project but surfaced when ProjectID=1002 accumulated sessions — the lngStepID=0 case matched all sessions globally. Fix is the standard TempVar pattern applied to a second dimension (ProjectID via join); George applied and exported, Claude verified logic and logged.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1040,1042,'2026-05-25','07:30:00','09:40:00',0,5,'Time(7) rendered as text in Access linked tables — display worked but editing failed; root cause discovered during YouTube demo recording. Migration required index drop before ALTER; SSE distro required export/import pattern to bypass first-run connection dialog blocking MCP design-view operations. Two standing rules documented: physical bound controls for PK/FK VBA references (Design-Patterns 1.13), and full CRUD test protocol (not display-only).');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1045,1044,'2026-05-27','05:00:00','07:45:00',0,2,'Recorded across five sessions on three days (5/25–5/27/2026); two revision cycles — at end of 5/25 morning session and end of 5/26 morning session; final launch video (full story + AI team development insights) planned as a separate future deliverable.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1046,1045,'2026-05-27','07:30:00','09:00:00',0,2,'Grill session identified data-first framing (not announcement) as the right approach for skeptical Access developers; context reset count confirmed at 6 (build phase only, 5 launch-phase resets excluded); wannabe-inclusive language adjusted for Reddit; YouTube description added as fifth version.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1050,1051,'2026-05-27','09:00:00','10:30:00',0,0,'Posted and edited announcements');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1051,1045,'2026-05-27','09:00:00','09:20:00',0,0,'Requested announcements tailored to different venues: X, YouTube, LinkedIn, Forums, Reddit');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1052,1017,'2026-05-27','08:00:00','08:15:00',0,0,'Gave Claude 5 forums/online sources types');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1053,1052,'2026-05-27','13:13:00','13:20:00',0,0,'Only two viable public Access community forums remain active; posted to all available platforms. Announcement phase complete; Project 1003 closed.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1063,1065,'2026-05-29','07:06:00','09:13:00',0,2,'Script developed through iterative grill session; PM behavior change angle added as primary narrative thread distinct from the demo video; data from tblSession (6 context resets, 72 revision cycles, 30 sessions) used to anchor all three PM points; Tom confirmed consent for email on camera.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1064,1066,'2026-05-30','09:00:00','19:00:00',0,4,'Solo recording and editing session -- 10 hours total; four revision cycles during recording and editing. Video live at https://youtu.be/dC-TJrjU7hA');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1065,1067,'2026-05-31','07:30:00','09:15:00',0,0,'Session open. Writing blog post from PM lessons learned script.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1066,1067,'2026-05-31','09:55:00','11:36:00',0,2,'Images double-resized (25% then 50%) and lost fidelity; recreate originals as PNG at ~50% capture size. Tom confirmed as top-tier Access developer; memory updated.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1067,1069,'2026-05-31','15:51:00','16:27:00',0,0,'Identified duplicate session 1049 (dropped) and corrected invalid StepTypeID on StepID 1050 before adding to seed data; Access MCP used for direct BE inserts; DevPulse_BE.accdb binary re-uploaded to GitHub Releases by George.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1068,1070,'2026-06-01','09:21:00','10:05:00',0,0,'Script approved after VS Code edit by George plus three targeted Claude fixes. Key addition: honest disclosure about tblTrack being somewhat staged; full 10-tool stack slide (including JoyfullService VCS) added to closing. StepID=1068 renamed to filming-and-upload scope per PM lessons split pattern.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (1069,1071,'2026-06-01','10:20:00','10:25:00',0,0,'File delete and schema drop ran in parallel; both succeeded. Albums schema and all 5 tables confirmed absent; MusicAlbums.accdb confirmed gone. Demo environment is back to pre-build state — ready for on-camera rebuild.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (3073,1068,'2026-06-05','07:00:00','16:00:00',1,2,'Recorded in 9 segments; final video ran 44:50 vs. planned 22-26 min. Two deliberate context resets during MusicAlbums demo to illustrate context management; two revision cycles.');
INSERT INTO tblSession (SessionID,StepID,SessionDate,SessionStartTime,SessionEndTime,ContextReset,RevisionCycles,Notes) VALUES (3074,3075,'2026-06-04','09:15:00','09:30:00',0,0,'Prompted by attendee who had read the blog post. Context reset explained via web-app state-loss analogy; close process compared to preserving session state between screens. GitHub repo shared with group.');
SET IDENTITY_INSERT tblSession OFF;

SET IDENTITY_INSERT tblSessionTool ON;
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1074,1034,1002,1);
INSERT INTO tblSessionTool (SessionToolID,SessionID,ToolTypeID,UseCount) VALUES (1075,1034,4,1);
SET IDENTITY_INSERT tblSessionTool OFF;


-- ============================================================
-- Setup complete.
-- Next step: connect the Access front end (DevPulseSSE.accdb)
-- via ODBC using frmSetupConnection, or open DevPulseACE.accdb
-- for the Access backend edition.
-- See README.md for full setup instructions.