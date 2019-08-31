USE FileOK
GO

DROP TABLE IF EXISTS dbo.FilesInventory
GO

CREATE TABLE dbo.FilesInventory(
	pk_id int IDENTITY(1,1) NOT NULL,
	DirectoryName varchar(2048) NOT NULL,
	ParentDirectoryName varchar(2048) NULL,
	DirectoryDateCreatedUtc date NULL,
	DirectoryDateLastAccessUtc date NULL,
	DirectoryDateLastWriteUtc date NULL,
	FileNamefull varchar(2048) NULL,
	FileExtension varchar(2048) NULL,
	FileSize varchar(2048) NULL,
	FileDateCreatedUtc date NULL,
	FileDateLastAccessUtc date NULL,
	FileDateLastWriteUtc date NULL,
	FileHashCode varchar(2048) NULL,
	FileReadOnly bit NULL,
PRIMARY KEY CLUSTERED 
(
	pk_id ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON PRIMARY
) ON PRIMARY
GO


