CREATE TABLE [USysLinkedTables] (
  [TableName] VARCHAR (100) CONSTRAINT [PK_USysLinkedTables] PRIMARY KEY UNIQUE NOT NULL,
  [SourceName] VARCHAR (100),
  [SortOrder] LONG
)
