CREATE TABLE [tmpl_tlkpOutputType] (
  [OutputTypeID] LONG CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [OutputTypeName] VARCHAR (220),
  [SortOrder] LONG,
  [InactiveDate] DATETIME
)
