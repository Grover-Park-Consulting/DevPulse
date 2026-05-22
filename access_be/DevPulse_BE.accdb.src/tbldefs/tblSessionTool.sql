CREATE TABLE [tblSessionTool] (
  [SessionToolID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SessionID] LONG CONSTRAINT [FK_tblSessionTool_Session] REFERENCES [tblSession] ([SessionID]),
  [ToolTypeID] LONG CONSTRAINT [FK_tblSessionTool_Tool] REFERENCES [tlkpTool] ([ToolTypeID]),
  [UseCount] LONG
)
