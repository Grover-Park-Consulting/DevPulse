CREATE TABLE [tlkpProjectArea] (
  [ProjectAreaID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ProjectAreaName] VARCHAR (100),
  [SortOrder] SHORT,
  [InactiveDate] DATETIME
)
