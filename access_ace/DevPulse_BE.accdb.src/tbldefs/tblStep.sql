CREATE TABLE [tblStep] (
  [StepID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ProjectID] LONG CONSTRAINT [FK_tblStep_Project] REFERENCES [tblProject] ([ProjectID]),
  [StepName] VARCHAR (200),
  [StepDescription] LONGTEXT,
  [StepTypeID] LONG CONSTRAINT [FK_tblStep_StepType] REFERENCES [tlkpStepType] ([StepTypeID]),
  [StatusID] LONG CONSTRAINT [FK_tblStep_Status] REFERENCES [tlkpStatus] ([StatusID]),
  [PredecessorStepID] LONG,
  [StartDate] DATETIME,
  [CompletionDate] DATETIME,
  [Notes] LONGTEXT,
  [CreatedDate] DATETIME,
  [CreatedBy] VARCHAR (100),
  [ModifiedDate] DATETIME,
  [ModifiedBy] VARCHAR (100)
)
