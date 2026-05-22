CREATE TABLE [tblLookupCacheConfig] (
  [CacheConfigID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [TemplateName] VARCHAR (100),
  [ODBCTableName] VARCHAR (100),
  [FieldList] VARCHAR (255),
  [SortOrder] LONG
)
