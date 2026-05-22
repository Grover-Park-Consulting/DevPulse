CREATE TABLE [tblOutput] (
  [OutputID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [StepID] LONG CONSTRAINT [FK_tblOutput_Step] REFERENCES [tblStep] ([StepID]),
  [OutputTypeID] LONG CONSTRAINT [FK_tblOutput_OutputType] REFERENCES [tlkpOutputType] ([OutputTypeID]),
  [OutputCount] LONG,
  [OutputDescription] LONGTEXT
)
