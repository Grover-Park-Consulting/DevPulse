CREATE TABLE [tlkpTool] (
  [ToolTypeID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ToolTypeName] VARCHAR (100),
  [SortOrder] SHORT,
  [InactiveDate] DATETIME
)
