/*********************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Ivan Muñoz
  Stored Procedure: spFormulasCMasivas
\*********************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spFormulasCMasivas]    Script Date: 13/02/2019 06:02:25 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spFormulasCMasivas')
  BEGIN
    DROP PROCEDURE [dbo].[spFormulasCMasivas]
  END
GO

/****** Object:  StoredProcedure [dbo].[spFormulasCMasivas]    Script Date: 13/02/2019 06:02:25 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spFormulasCMasivas]
AS
BEGIN
  DECLARE @JSON VARCHAR(MAX) = '['
  SELECT
   @JSON = @JSON +
   '{"Posicion":"' + CAST(CAST(REPLACE(REPLACE(SUBSTRING(Formula, 0, 4), '[0', ' '), '[', '') AS INT) - 1 AS VARCHAR(3)) +
   '", "Anexo":"' + 'ANEXO' + REPLACE(REPLACE(SUBSTRING(Formula, 0, 4), '[0', ' '), '[', ' ') +
   '", "Indice":"' + SUBSTRING(Formula, 5, 14) +
   '", "Celda":"' + REPLACE(REPLACE(SUBSTRING(Formula, 20, 3), ']=',''), ']', '') +
   '", "Formula":"' + REPLACE(REPLACE(REPLACE(SUBSTRING(Formula, CHARINDEX('=', Formula) + 1, LEN(Formula)), '[' + REPLACE(SUBSTRING(Formula, 0, 4), '[', '') + ',', ''), ']', ''), REPLACE(REPLACE(SUBSTRING(Formula, 19, 3), ']=',''), ']', ''), '') + '"},'
  FROM tbl_Comprobaciones
  WHERE IdTipoPlantilla = 1
  AND CAST(REPLACE(REPLACE(SUBSTRING(Formula, 0, 4), '[0', ' '), '[', '') AS INT) - 1 < 8

  SET	@JSON	=	SUBSTRING(@JSON, 1, LEN(@JSON) -1) + ']'
  SELECT @JSON AS [Json]
END;
GO