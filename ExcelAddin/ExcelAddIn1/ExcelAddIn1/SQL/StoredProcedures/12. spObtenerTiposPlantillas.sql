/*****************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Ivan Muñoz
  Stored Procedure: spObtenerTiposPlantillas
\*****************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spObtenerTiposPlantillas]    Script Date: 13/02/2019 06:11:18 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spObtenerTiposPlantillas')
  BEGIN
    DROP PROCEDURE [dbo].[spObtenerTiposPlantillas]
  END
GO

/****** Object:  StoredProcedure [dbo].[spObtenerTiposPlantillas]    Script Date: 13/02/2019 06:11:18 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spObtenerTiposPlantillas]
AS
BEGIN
	SELECT		IdTipoPlantilla
				,Clave
				,Concepto
		FROM	dbo.tbl_TiposPlantillas
		FOR JSON PATH;
END;
GO
