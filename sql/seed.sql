-- ============================================================
--  BlazorWordAddIn  –  Seed Data
--  Replace UPNs below with real test accounts before running.
--  This script is idempotent – safe to run multiple times.
-- ============================================================

-- ── Helpers ───────────────────────────────────────────────────────────────────
SET NOCOUNT ON;
BEGIN TRANSACTION;

-- Test users
DECLARE @user1 NVARCHAR(320) = N'admin@M365x96996669.onmicrosoft.com';
DECLARE @user2 NVARCHAR(320) = N'AlexW@M365x96996669.OnMicrosoft.com';
DECLARE @user3 NVARCHAR(320) = N'AlexW@M365x96996669.OnMicrosoft.com'; -- no third test account; reuses AlexW

-- A tiny 10×10 red PNG in base64 (placeholder – replace with real images)
DECLARE @testImageB64 NVARCHAR(MAX) =
    N'iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAF0lEQVR42mP8' +
    N'z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg==';

-- ── Text Snippets ─────────────────────────────────────────────────────────────
-- alice's snippets
IF NOT EXISTS (SELECT 1 FROM dbo.TextSnippets WHERE OwnerUpn = @user1 AND Title = N'Monthly Status Report Header')
    INSERT INTO dbo.TextSnippets (OwnerUpn, Title, BodyText, Tags)
    VALUES (@user1, N'Monthly Status Report Header',
        N'Monthly Status Report' + CHAR(13)+CHAR(10) +
        N'Period: [MONTH YEAR]' + CHAR(13)+CHAR(10) +
        N'Prepared by: [NAME]' + CHAR(13)+CHAR(10) +
        N'Distribution: [LIST]',
        N'report,header,monthly');

IF NOT EXISTS (SELECT 1 FROM dbo.TextSnippets WHERE OwnerUpn = @user1 AND Title = N'Action Item Template')
    INSERT INTO dbo.TextSnippets (OwnerUpn, Title, BodyText, Tags)
    VALUES (@user1, N'Action Item Template',
        N'ACTION ITEM' + CHAR(13)+CHAR(10) +
        N'Owner: [NAME]' + CHAR(13)+CHAR(10) +
        N'Due Date: [DATE]' + CHAR(13)+CHAR(10) +
        N'Priority: [High / Medium / Low]' + CHAR(13)+CHAR(10) +
        N'Description: [DETAILS]',
        N'action,template');

IF NOT EXISTS (SELECT 1 FROM dbo.TextSnippets WHERE OwnerUpn = @user1 AND Title = N'Classification Banner – UNCLASSIFIED')
    INSERT INTO dbo.TextSnippets (OwnerUpn, Title, BodyText, Tags)
    VALUES (@user1, N'Classification Banner – UNCLASSIFIED',
        N'UNCLASSIFIED // FOR OFFICIAL USE ONLY',
        N'classification,banner,fouo');

-- bob's snippets
IF NOT EXISTS (SELECT 1 FROM dbo.TextSnippets WHERE OwnerUpn = @user2 AND Title = N'Meeting Minutes Opening')
    INSERT INTO dbo.TextSnippets (OwnerUpn, Title, BodyText, Tags)
    VALUES (@user2, N'Meeting Minutes Opening',
        N'Meeting Minutes' + CHAR(13)+CHAR(10) +
        N'Date: [DATE]  Time: [TIME]  Location: [LOCATION / VTC Link]' + CHAR(13)+CHAR(10) +
        N'Attendees: [LIST]' + CHAR(13)+CHAR(10) +
        N'Facilitator: [NAME]' + CHAR(13)+CHAR(10) +
        N'Note Taker: [NAME]',
        N'meeting,minutes');

IF NOT EXISTS (SELECT 1 FROM dbo.TextSnippets WHERE OwnerUpn = @user2 AND Title = N'Risk Statement')
    INSERT INTO dbo.TextSnippets (OwnerUpn, Title, BodyText, Tags)
    VALUES (@user2, N'Risk Statement',
        N'Risk ID: [ID]' + CHAR(13)+CHAR(10) +
        N'Description: [RISK DESCRIPTION]' + CHAR(13)+CHAR(10) +
        N'Likelihood: [High / Medium / Low]' + CHAR(13)+CHAR(10) +
        N'Impact: [High / Medium / Low]' + CHAR(13)+CHAR(10) +
        N'Mitigation: [MITIGATION PLAN]',
        N'risk,template');

-- carol's snippets
IF NOT EXISTS (SELECT 1 FROM dbo.TextSnippets WHERE OwnerUpn = @user3 AND Title = N'Executive Summary Opener')
    INSERT INTO dbo.TextSnippets (OwnerUpn, Title, BodyText, Tags)
    VALUES (@user3, N'Executive Summary Opener',
        N'Executive Summary' + CHAR(13)+CHAR(10) +
        N'This document provides [a brief statement of topic] for the period of [DATE RANGE]. ' +
        N'Key findings are summarized below.',
        N'executive,summary');

-- ── Image Snippets ────────────────────────────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM dbo.ImageSnippets WHERE OwnerUpn = @user1 AND Title = N'Org Chart Placeholder')
    INSERT INTO dbo.ImageSnippets (OwnerUpn, Title, ImageBase64, MimeType, Tags)
    VALUES (@user1, N'Org Chart Placeholder', @testImageB64, N'image/png', N'org,chart');

IF NOT EXISTS (SELECT 1 FROM dbo.ImageSnippets WHERE OwnerUpn = @user2 AND Title = N'Agency Logo Placeholder')
    INSERT INTO dbo.ImageSnippets (OwnerUpn, Title, ImageBase64, MimeType, Tags)
    VALUES (@user2, N'Agency Logo Placeholder', @testImageB64, N'image/png', N'logo,branding');

COMMIT TRANSACTION;
PRINT 'Seed data inserted successfully.';
