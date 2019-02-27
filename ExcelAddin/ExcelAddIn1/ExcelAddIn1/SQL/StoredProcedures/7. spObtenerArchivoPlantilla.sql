/*********************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Ivan Muñoz
  Stored Procedure: spObtenerArchivoPlantilla
\*********************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spObtenerArchivoPlantilla]    Script Date: 13/02/2019 05:59:25 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spObtenerArchivoPlantilla')
  BEGIN
    DROP PROCEDURE [dbo].[spObtenerArchivoPlantilla]
  END
GO

/****** Object:  StoredProcedure [dbo].[spObtenerArchivoPlantilla]    Script Date: 13/02/2019 05:59:25 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spObtenerArchivoPlantilla]
	@pIdPlantilla		INTEGER
AS
BEGIN
	SELECT		ArchivoPlantilla
		FROM	dbo.tbl_Plantillas
		WHERE	IdPlantilla	=	@pIdPlantilla
END;
GO