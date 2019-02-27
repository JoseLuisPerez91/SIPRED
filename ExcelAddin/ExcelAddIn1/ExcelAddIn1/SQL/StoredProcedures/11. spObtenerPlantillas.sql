/*****************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Ivan Muñoz
  Stored Procedure: spObtenerPlantillas
\*****************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spObtenerPlantillas]    Script Date: 13/02/2019 06:09:10 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spObtenerPlantillas')
  BEGIN
    DROP PROCEDURE [dbo].[spObtenerPlantillas]
  END
GO

/****** Object:  StoredProcedure [dbo].[spObtenerPlantillas]    Script Date: 13/02/2019 06:09:10 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spObtenerPlantillas]
AS
BEGIN
	SELECT		IdPlantilla
				,IdTipoPlantilla
				,Anio
				,Nombre
				,Usuario_Modifico AS Usuario
		FROM	dbo.tbl_Plantillas
		WHERE	Activo		=	1
		FOR JSON PATH;
END;
GO
