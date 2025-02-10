IF (EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND  TABLE_NAME = 'WEB_MENU'))
BEGIN
    DROP TABLE WEB_MENU
END
CREATE TABLE [dbo].[WEB_MENU](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[nome] [nvarchar](255) NOT NULL,
	[path] [nvarchar](255) NULL,
	[icon] [nvarchar](255) NULL,
	[parent] [int] NOT NULL,
	[admin] [bit] NULL,
	[cliente] [bit] NULL,
	[ordine] [int] NULL,
	[note] [ntext] NOT NULL,
 CONSTRAINT [PK_WEB_MENU] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

ALTER TABLE [dbo].[WEB_MENU] ADD  CONSTRAINT [DF_WEB_MENU_parent]  DEFAULT ((-1)) FOR [parent]
GO

IF (EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND  TABLE_NAME = 'WEB_SINOTTICO'))
BEGIN
    DROP TABLE WEB_SINOTTICO
END
CREATE TABLE [dbo].[WEB_SINOTTICO](
	[verde] [nvarchar](255) NOT NULL,
	[giallo] [nvarchar](255) NOT NULL,
	[rosso] [nvarchar](255) NOT NULL,
	[id] [nvarchar](255) NOT NULL,
	[tipo] [int] NOT NULL,
	[unita_misura] [nvarchar](255) NOT NULL,
	[tag_rif] [nvarchar](255) NOT NULL,
	[colore_default] [nvarchar](255) NOT NULL,
	[resizable] [bit] NOT NULL,
	[indice] [int] IDENTITY(1,1) NOT NULL,
 CONSTRAINT [PK_WEB_SINOTTICO] PRIMARY KEY CLUSTERED 
(
	[indice] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[WEB_SINOTTICO] ADD  CONSTRAINT [DF_WEB_SINOTTICO_NEW_colore_default]  DEFAULT (N'#9b9b9b') FOR [colore_default]
GO

ALTER TABLE [dbo].[WEB_SINOTTICO] ADD  CONSTRAINT [DF_WEB_SINOTTICO_resizable]  DEFAULT ((0)) FOR [resizable]
GO

IF (EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND  TABLE_NAME = 'WEB_UTENTI'))
BEGIN
    DROP TABLE WEB_UTENTI
END

CREATE TABLE [dbo].[WEB_UTENTI](
	[username] [nvarchar](255) NOT NULL,
	[password] [nvarchar](255) NOT NULL,
	[tipo] [int] NOT NULL
) ON [PRIMARY]

GO
--Dati di default per la tabella menu
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Diario', '', '<i class="far fa-calendar-alt"></i>', -1, 1, 1, 5, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Tabella', '../Pagine/DiarioTabella', '<i class="fas fa-table"></i>', 1, 1, 1, 1, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Grafico', '../Pagine/DiarioGrafico', '<i class="fas fa-chart-line"></i>', 1, 1, 1, 2, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('allarmi_stati', '../Pagine/AllarmiStati', '<i class="fas fa-bell"></i>', -1, 1, 1, 3, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Sinottico', '../Pagine/Sinottico', '<i class="fas fa-project-diagram"></i>', -1, 1, 1, 2, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Config.', '', '<i class="fas fa-cog"></i>', -1, 1, 1, 4, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Report', '../Pagine/Report', '<i class="far fa-file-excel"></i>', -1, 1, 1, 6, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Menu', '../Pagine/SettingMenu', '<i class="fas fa-bars"></i>', 6, 1, 0, 3, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Utenti', '../Pagine/SettingUtente', '<i class="fas fa-users"></i>', 6, 1, 0, 4, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Sinottico', '../Pagine/SettingSinottico', '<i class="fas fa-project-diagram"></i>', 6, 1, 0, 5, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Generali', '../Pagine/Setting', '<i class="fas fa-cog"></i>', 6, 1, 0, 6, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Misure', '../Pagine/Misure', '<i class="fas fa-align-left"></i>', -1, 1, 1, 1, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Soglie', '../Pagine/SettingSoglie', '<i class="fas fa-sliders-h"></i>', 6, 1, 1, 1, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('QAL2', '../Pagine/SettingQUAL2', '<i class="far fa-edit"></i>', 6, 1, 1, 2, '')
INSERT INTO [WEB_MENU] ([nome], [path], [icon], [parent], [admin], [cliente], [ordine], [note]) VALUES ('Languages', '../Pagine/SettingLanguages', '<i class="fas fa-language"></i>', 6, 1, 0, 9, '')
--Dati di default per la tabella utenti
INSERT INTO [WEB_UTENTI] ([username] ,[password] ,[tipo]) VALUES ('sme', '06bf47c73d0d12aa04acf1157a511aab', 1)
INSERT INTO [WEB_UTENTI] ([username] ,[password] ,[tipo]) VALUES ('bf', '69854fd9506319b96e1693832159815a', 0)

