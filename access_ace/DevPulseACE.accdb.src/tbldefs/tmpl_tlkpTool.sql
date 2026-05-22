CREATE TABLE [tmpl_tlkpTool] (
  [ToolTypeID] LONG CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ToolTypeName] VARCHAR (220),
  [SortOrder] LONG,
  [InactiveDate] DATETIME
)
