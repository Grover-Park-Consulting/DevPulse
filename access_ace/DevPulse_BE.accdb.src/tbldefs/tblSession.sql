CREATE TABLE [tblSession] (
  [SessionID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [StepID] LONG CONSTRAINT [FK_tblSession_Step] REFERENCES [tblStep] ([StepID]),
  [SessionDate] DATETIME,
  [SessionStartTime] DATETIME,
  [SessionEndTime] DATETIME,
  [ContextReset] BIT,
  [RevisionCycles] LONG,
  [Notes] LONGTEXT,
  [CreatedDate] DATETIME,
  [CreatedBy] VARCHAR (100),
  [ModifiedDate] DATETIME,
  [ModifiedBy] VARCHAR (100)
)
