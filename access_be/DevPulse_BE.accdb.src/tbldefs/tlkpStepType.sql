CREATE TABLE [tlkpStepType] (
  [StepTypeID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [StepTypeName] VARCHAR (100),
  [SortOrder] SHORT,
  [InactiveDate] DATETIME
)
