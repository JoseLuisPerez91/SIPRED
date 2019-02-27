/*****************************************\
  Proyect  : Deloitte ExcelAddIn
  Developer: Ing. Ivan Muñoz
  Action   : Create Table tbl_TiposPlantillas
\*****************************************/
USE [DSAT]
GO
/****** Object:  Table [dbo].[tbl_Plantillas] ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'tbl_TiposPlantillas')
  BEGIN
    DROP TABLE [dbo].[tbl_TiposPlantillas]
  END
GO
/****** Object:  Table [dbo].[tbl_TiposPlantillas] ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_TiposPlantillas](
	[IdTipoPlantilla] [int] NOT NULL,
	[Clave] [varchar](50) NOT NULL,
	[Concepto] [varchar](150) NOT NULL,
 CONSTRAINT [PK_tbl_TiposPlantillas] PRIMARY KEY CLUSTERED 
(
	[IdTipoPlantilla] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****************************************\
  Proyect  : Deloitte ExcelAddIn
  Developer: Ing. Ivan Muñoz
  Action   : Insert Data tbl_TiposPlantillas
\****************************************/
INSERT INTO tbl_TiposPlantillas(IdTipoPlantilla, Clave, Concepto) VALUES(1, 'SIPRED', 'Estados Financieros General')
INSERT INTO tbl_TiposPlantillas(IdTipoPlantilla, Clave, Concepto) VALUES(2, 'ISSIF', 'Personas Morales en general')
/************************************\
  Proyect  : Deloitte ExcelAddIn
  Developer: Ing. Ivan Muñoz
  Action   : Select Table tbl_TiposPlantillas
\************************************/
SELECT IdTipoPlantilla, Clave, Concepto
FROM tbl_TiposPlantillas
ORDER BY IdTipoPlantilla
GO
