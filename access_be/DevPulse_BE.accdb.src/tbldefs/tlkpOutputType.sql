CREATE TABLE [tlkpOutputType] (
  [OutputTypeID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [OutputTypeName] VARCHAR (100),
  [SortOrder] SHORT,
  [InactiveDate] DATETIME
)
