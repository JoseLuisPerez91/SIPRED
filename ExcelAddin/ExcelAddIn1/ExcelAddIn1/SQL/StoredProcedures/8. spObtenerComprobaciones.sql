/*********************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Ivan Muñoz
  Stored Procedure: spObtenerComprobaciones
\*********************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spObtenerComprobaciones]    Script Date: 13/02/2019 06:02:25 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spObtenerComprobaciones')
  BEGIN
    DROP PROCEDURE [dbo].[spObtenerComprobaciones]
  END
GO

/****** Object:  StoredProcedure [dbo].[spObtenerComprobaciones]    Script Date: 13/02/2019 06:02:25 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spObtenerComprobaciones]
AS
BEGIN
	DECLARE
		@JSON	VARCHAR(MAX)		=	'[';

	SELECT		@JSON = @JSON + '{"IdComprobacion": ' + CONVERT(VARCHAR, IdComprobacion) + ', "IdTipoPlantilla": ' + CONVERT(VARCHAR, IdTipoPlantilla) + ', ' + 
				'"Concepto": "' + RTRIM(LTRIM(REPLACE([Concepto], '"', '\"'))) + '", "Formula": "' + RTRIM(LTRIM(Formula)) + '", "Condicion": "' + RTRIM(LTRIM(Condicion)) + '","Nota":"'+Nota+'","AdmiteCambios":"'+CONVERT(VARCHAR, AdmiteCambios)+'"},'
		FROM	dbo.tbl_Comprobaciones
		--FOR JSON PATH;

	SET	@JSON	=	SUBSTRING(@JSON, 1, LEN(@JSON) -1) + ']';
	
	SELECT @JSON AS [Json]; 
END;
GO
