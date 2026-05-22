CREATE TABLE [tblProject] (
  [ProjectID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ProjectName] VARCHAR (200),
  [ProjectDescription] LONGTEXT,
  [ProjectAreaID] LONG CONSTRAINT [FK_tblProject_Area] REFERENCES [tlkpProjectArea] ([ProjectAreaID]),
  [StatusID] LONG CONSTRAINT [FK_tblProject_Status] REFERENCES [tlkpStatus] ([StatusID]),
  [StartDate] DATETIME,
  [CompletionDate] DATETIME,
  [Notes] LONGTEXT,
  [CreatedDate] DATETIME,
  [CreatedBy] VARCHAR (100),
  [ModifiedDate] DATETIME,
  [ModifiedBy] VARCHAR (100)
)
