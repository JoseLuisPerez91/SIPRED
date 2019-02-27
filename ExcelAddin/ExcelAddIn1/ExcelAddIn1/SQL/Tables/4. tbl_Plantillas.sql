/*****************************************\
  Proyect  : Deloitte ExcelAddIn
  Developer: Ing. Ivan Muñoz
  Action   : Create Table tbl_Plantillas
\*****************************************/
USE [DSAT]
GO
/*IF  EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'tbl_Plantillas')
  BEGIN
    ALTER TABLE [dbo].[tbl_Plantillas] DROP CONSTRAINT [FK_tbl_Plantillas_tbl_TiposPlantillas]
  END
GO*/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'DF_tbl_Plantillas_Fecha_Modificacion')
  BEGIN
    ALTER TABLE [dbo].[tbl_Plantillas] DROP CONSTRAINT [DF_tbl_Plantillas_Fecha_Modificacion]
  END
GO
/****** Object:  Table [dbo].[tbl_Plantillas] ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'tbl_Plantillas')
  BEGIN
    DROP TABLE [dbo].[tbl_Plantillas]
  END
GO
/****** Object:  Table [dbo].[tbl_Plantillas]    Script Date: 18/01/2019 10:12:35 p. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_Plantillas](
	[IdPlantilla] [int] NOT NULL,
	[IdTipoPlantilla] [int] NOT NULL,
	[Anio] [int] NOT NULL,
	[Nombre] [varchar](1500) NOT NULL,
	[ArchivoPlantilla] [varbinary](max) NULL,
	[Activo] [bit] NOT NULL,
	[Fecha_Modificacion] [datetime] NOT NULL,
	[Usuario_Modifico] [varchar](150) NOT NULL,
 CONSTRAINT [PK_tbl_Plantillas] PRIMARY KEY CLUSTERED 
(
	[IdPlantilla] ASC,
	[IdTipoPlantilla] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[tbl_Plantillas] ADD CONSTRAINT [DF_tbl_Plantillas_Fecha_Modificacion]  DEFAULT (getdate()) FOR [Fecha_Modificacion]
GO
/*ALTER TABLE [dbo].[tbl_Plantillas]  WITH CHECK ADD CONSTRAINT [FK_tbl_Plantillas_tbl_TiposPlantillas] FOREIGN KEY([IdTipoPlantilla])
REFERENCES [dbo].[tbl_TiposPlantillas] ([IdTipoPlantilla])
GO
ALTER TABLE [dbo].[tbl_Plantillas] CHECK CONSTRAINT [FK_tbl_Plantillas_tbl_TiposPlantillas]
GO*/
/****************************************\
  Proyect  : Deloitte ExcelAddIn
  Developer: Ing. Ivan Muñoz
  Action   : Insert Data tbl_Plantillas
\****************************************/
INSERT INTO tbl_Plantillas(IdPlantilla, IdTipoPlantilla, Anio, Nombre, ArchivoPlantilla, Activo, Fecha_Modificacion, Usuario_Modifico) VALUES(1, 1, 2019, 'SIPRED-EstadosFinancierosGeneral.xlsm', NULL, 0, '2019-01-18 21:19:52.977', 'app_sipred')
/************************************\
  Proyect  : Deloitte ExcelAddIn
  Developer: Ing. Ivan Muñoz
  Action   : Select Table tbl_TiposPlantillas
\************************************/
SELECT IdPlantilla, IdTipoPlantilla, Anio, Nombre, ArchivoPlantilla, Activo, Fecha_Modificacion, Usuario_Modifico
FROM tbl_Plantillas
ORDER BY IdPlantilla
GO


