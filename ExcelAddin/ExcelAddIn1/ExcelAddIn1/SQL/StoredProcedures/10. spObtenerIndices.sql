/*****************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Ivan Muñoz
  Stored Procedure: spObtenerIndices
\*****************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spObtenerIndices]    Script Date: 13/02/2019 06:06:46 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spObtenerIndices')
  BEGIN
    DROP PROCEDURE [dbo].[spObtenerIndices]
  END
GO

/****** Object:  StoredProcedure [dbo].[spObtenerIndices]    Script Date: 13/02/2019 06:06:46 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spObtenerIndices]
AS
BEGIN
  DECLARE @Table AS TABLE
  (
   Registro VARCHAR(MAX)
  )

  INSERT INTO @Table(Registro)
  VALUES('{"Subtotales":[{"Hoja":"ANEXO5","Columna":"C"},{"Hoja":"ANEXO5","Columna":"D"},{"Hoja":"ANEXO5","Columna":"J"},{"Hoja":"ANEXO5","Columna":"K"},{"Hoja":"ANEXO5","Columna":"L"},{"Hoja":"ANEXO5","Columna":"M"},{"Hoja":"ANEXO7","Columna":"C"},{"Hoja":"ANEXO7","Columna":"D"},{"Hoja":"ANEXO7","Columna":"J"},{"Hoja":"ANEXO7","Columna":"K"},{"Hoja":"ANEXO7","Columna":"Q"},{"Hoja":"ANEXO7","Columna":"R"},{"Hoja":"ANEXO7","Columna":"X"},{"Hoja":"ANEXO7","Columna":"Y"},{"Hoja":"ANEXO7","Columna":"AE"},{"Hoja":"ANEXO7","Columna":"AF"},{"Hoja":"ANEXO7","Columna":"AG"},{"Hoja":"ANEXO7","Columna":"AH"},{"Hoja":"ANEXO7","Columna":"AI"},{"Hoja":"ANEXO7","Columna":"AJ"},{"Hoja":"ANEXO6","Columna":"C"},{"Hoja":"ANEXO6","Columna":"D"},{"Hoja":"ANEXO8","Columna":"C"},{"Hoja":"ANEXO8","Columna":"D"},{"Hoja":"ANEXO8","Columna":"E"},{"Hoja":"ANEXO8","Columna":"F"},{"Hoja":"ANEXO8","Columna":"G"},{"Hoja":"ANEXO8","Columna":"H"},{"Hoja":"ANEXO2","Columna":"C"},{"Hoja":"ANEXO2","Columna":"D"},{"Hoja":"ANEXO2","Columna":"E"},{"Hoja":"ANEXO2","Columna":"F"},{"Hoja":"ANEXO2","Columna":"G"},{"Hoja":"ANEXO2","Columna":"H"},{"Hoja":"ANEXO10","Columna":"C"},{"Hoja":"ANEXO10","Columna":"D"},{"Hoja":"ANEXO10","Columna":"E"},{"Hoja":"ANEXO10","Columna":"F"},{"Hoja":"ANEXO3","Columna":"C"},{"Hoja":"ANEXO3","Columna":"D"},{"Hoja":"ANEXO3","Columna":"E"},{"Hoja":"ANEXO3","Columna":"F"},{"Hoja":"ANEXO3","Columna":"G"},{"Hoja":"ANEXO3","Columna":"H"},{"Hoja":"ANEXO3","Columna":"I"},{"Hoja":"ANEXO3","Columna":"J"},{"Hoja":"ANEXO3","Columna":"K"},{"Hoja":"ANEXO3","Columna":"L"},{"Hoja":"ANEXO3","Columna":"M"},{"Hoja":"ANEXO3","Columna":"N"},{"Hoja":"ANEXO3","Columna":"O"},{"Hoja":"ANEXO3","Columna":"P"},{"Hoja":"ANEXO3","Columna":"Q"},{"Hoja":"ANEXO3","Columna":"R"},{"Hoja":"ANEXO3","Columna":"S"},{"Hoja":"ANEXO3","Columna":"T"},{"Hoja":"ANEXO3","Columna":"U"},{"Hoja":"ANEXO1","Columna":"C"},{"Hoja":"ANEXO1","Columna":"D"},{"Hoja":"ANEXO20","Columna":"C"},{"Hoja":"ANEXO20","Columna":"D"},{"Hoja":"ANEXO20","Columna":"E"},{"Hoja":"ANEXO20","Columna":"F"},{"Hoja":"ANEXO20","Columna":"G"},{"Hoja":"ANEXO20","Columna":"H"},{"Hoja":"ANEXO4","Columna":"C"},{"Hoja":"ANEXO4","Columna":"D"},{"Hoja":"ANEXO11","Columna":"C"},{"Hoja":"ANEXO15","Columna":"C"},{"Hoja":"ANEXO9","Columna":"C"},{"Hoja":"ANEXO9","Columna":"E"},{"Hoja":"ANEXO9","Columna":"F"},{"Hoja":"ANEXO9","Columna":"G"},{"Hoja":"ANEXO9","Columna":"H"}],"Conceptos":[{"Descripcion":"OTRO","Caracteres":4},{"Descripcion":"(","Caracteres":1},{"Descripcion":"*","Caracteres":1},{"Descripcion":"OTRA","Caracteres":4}]}')

	SELECT Registro
	FROM @Table
		--FOR JSON PATH;
END;
GO
