CREATE TABLE [tblAppConfig] (
  [ConfigName] VARCHAR (100) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ConfigValue] LONGTEXT,
  [ConfigDescription] LONGTEXT
)
