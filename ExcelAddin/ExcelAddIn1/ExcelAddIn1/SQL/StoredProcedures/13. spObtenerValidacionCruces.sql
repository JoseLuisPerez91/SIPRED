/*********************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Ivan Muñoz
  Stored Procedure: spObtenerValidacionCruces
\*********************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spObtenerValidacionCruces]    Script Date: 13/02/2019 06:12:56 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spObtenerValidacionCruces')
  BEGIN
    DROP PROCEDURE [dbo].[spObtenerValidacionCruces]
  END
GO

/****** Object:  StoredProcedure [dbo].[spObtenerValidacionCruces]    Script Date: 13/02/2019 06:12:56 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spObtenerValidacionCruces]
AS
BEGIN
  DECLARE @Table AS TABLE
  (
   Registro VARCHAR(MAX)
  )

  INSERT INTO @Table(Registro)
  VALUES('[{"Hoja":"Generales","Indice":"01D021000"},{"Hoja":"Generales","Indice":"01D022000"},{"Hoja":"Generales","Indice":"01D023000"},{"Hoja":"Generales","Indice":"01D024000"},{"Hoja":"Generales","Indice":"01D025000"},{"Hoja":"Generales","Indice":"01D027000"},{"Hoja":"Generales","Indice":"01D028000"},{"Hoja":"Generales","Indice":"01D034000"},{"Hoja":"Generales","Indice":"01D037000"},{"Hoja":"Generales","Indice":"01D038000"},{"Hoja":"Generales","Indice":"01D060000"},{"Hoja":"Generales","Indice":"01D061000"},{"Hoja":"Generales","Indice":"01D062000"},{"Hoja":"Generales","Indice":"01D063000"},{"Hoja":"Generales","Indice":"01D064000"}]')

	SELECT Registro
	FROM @Table
		--FOR JSON PATH;
END;
GO
