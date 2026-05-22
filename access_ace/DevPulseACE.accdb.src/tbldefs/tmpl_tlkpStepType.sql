CREATE TABLE [tmpl_tlkpStepType] (
  [StepTypeID] LONG CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [StepTypeName] VARCHAR (220),
  [SortOrder] LONG,
  [InactiveDate] DATETIME
)
