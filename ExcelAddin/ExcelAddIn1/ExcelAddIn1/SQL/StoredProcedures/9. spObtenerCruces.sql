/*****************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Ivan Muñoz
  Stored Procedure: spObtenerCruces
\*****************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spObtenerCruces]    Script Date: 13/02/2019 06:04:43 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spObtenerCruces')
  BEGIN
    DROP PROCEDURE [dbo].[spObtenerCruces]
  END
GO

/****** Object:  StoredProcedure [dbo].[spObtenerCruces]    Script Date: 13/02/2019 06:04:43 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spObtenerCruces]
AS
BEGIN
	DECLARE
		@JSON	VARCHAR(MAX)		=	'[';

	SELECT		@JSON = @JSON + '{"IdCruce": ' + CONVERT(VARCHAR, IdCruce) + ', "IdTipoPlantilla": ' + CONVERT(VARCHAR, IdTipoPlantilla) + ', ' + 
				'"Concepto": "' + RTRIM(LTRIM(REPLACE([Concepto], '"', '\"'))) + '", "Formula": "' + RTRIM(LTRIM(Formula)) + '", "Condicion": "' + REPLACE(RTRIM(LTRIM(Condicion)), '"', '\\"') +
				'", "Nota": "' + Nota + '", "LecturaImportes": "' + CAST(LecturaImportes AS VARCHAR(2)) + '"},'
		FROM	dbo.tbl_Cruces
		--FOR JSON PATH;

	SET	@JSON	=	SUBSTRING(@JSON, 1, LEN(@JSON) -1) + ']';
	
	SELECT @JSON AS [Json]; 
END;
GO
