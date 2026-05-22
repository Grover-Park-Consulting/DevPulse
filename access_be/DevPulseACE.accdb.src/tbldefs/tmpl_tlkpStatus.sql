CREATE TABLE [tmpl_tlkpStatus] (
  [StatusID] LONG CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [StatusName] VARCHAR (220),
  [SortOrder] LONG,
  [InactiveDate] DATETIME
)
