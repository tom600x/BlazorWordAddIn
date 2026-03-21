-- ============================================================
--  BlazorWordAddIn  –  Schema
--  Target: Azure SQL Database (also runs on LocalDB for dev)
--  Run this script as the Entra admin of the SQL Server.
-- ============================================================

-- ── TextSnippets ──────────────────────────────────────────────────────────────
IF OBJECT_ID(N'dbo.TextSnippets', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TextSnippets (
        Id          INT              NOT NULL IDENTITY(1,1),
        OwnerUpn    NVARCHAR(320)    NOT NULL,
        Title       NVARCHAR(200)    NOT NULL,
        BodyText    NVARCHAR(MAX)    NOT NULL,
        Tags        NVARCHAR(500)        NULL,
        UpdatedUtc  DATETIME2(0)     NOT NULL CONSTRAINT DF_TextSnippets_UpdatedUtc DEFAULT SYSUTCDATETIME(),

        CONSTRAINT PK_TextSnippets PRIMARY KEY CLUSTERED (Id)
    );

    CREATE NONCLUSTERED INDEX IX_TextSnippets_OwnerUpn
        ON dbo.TextSnippets (OwnerUpn);
END
GO

-- ── ImageSnippets ─────────────────────────────────────────────────────────────
IF OBJECT_ID(N'dbo.ImageSnippets', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ImageSnippets (
        Id           INT              NOT NULL IDENTITY(1,1),
        OwnerUpn     NVARCHAR(320)    NOT NULL,
        Title        NVARCHAR(200)    NOT NULL,
        -- Base64-encoded image; stored as NVARCHAR(MAX).
        -- Consider Azure Blob Storage for large images in production.
        ImageBase64  NVARCHAR(MAX)    NOT NULL,
        MimeType     NVARCHAR(50)     NOT NULL CONSTRAINT DF_ImageSnippets_MimeType DEFAULT N'image/png',
        Tags         NVARCHAR(500)        NULL,
        UpdatedUtc   DATETIME2(0)     NOT NULL CONSTRAINT DF_ImageSnippets_UpdatedUtc DEFAULT SYSUTCDATETIME(),

        CONSTRAINT PK_ImageSnippets PRIMARY KEY CLUSTERED (Id)
    );

    CREATE NONCLUSTERED INDEX IX_ImageSnippets_OwnerUpn
        ON dbo.ImageSnippets (OwnerUpn);
END
GO

-- ── AllowedUsers ──────────────────────────────────────────────────────────────
-- Grants a secondary user read access to all of an owner's snippets.
IF OBJECT_ID(N'dbo.AllowedUsers', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AllowedUsers (
        OwnerUpn    NVARCHAR(320)    NOT NULL,
        AllowedUpn  NVARCHAR(320)    NOT NULL,
        GrantedUtc  DATETIME2(0)     NOT NULL CONSTRAINT DF_AllowedUsers_GrantedUtc DEFAULT SYSUTCDATETIME(),

        CONSTRAINT PK_AllowedUsers PRIMARY KEY CLUSTERED (OwnerUpn, AllowedUpn)
    );

    CREATE NONCLUSTERED INDEX IX_AllowedUsers_AllowedUpn
        ON dbo.AllowedUsers (AllowedUpn);
END
GO
