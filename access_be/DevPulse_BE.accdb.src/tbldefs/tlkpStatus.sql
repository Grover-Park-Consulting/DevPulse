CREATE TABLE [tlkpStatus] (
  [StatusID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [StatusName] VARCHAR (100),
  [SortOrder] SHORT,
  [InactiveDate] DATETIME
)
