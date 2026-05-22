CREATE TABLE [tmpl_tlkpProjectArea] (
  [ProjectAreaID] LONG CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ProjectAreaName] VARCHAR (220),
  [SortOrder] LONG,
  [InactiveDate] DATETIME
)
