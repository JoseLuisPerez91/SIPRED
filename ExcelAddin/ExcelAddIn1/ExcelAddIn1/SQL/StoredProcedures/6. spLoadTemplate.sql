/*****************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Ivan Muñoz
  Stored Procedure: spLoadTemplate
\*****************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spLoadTemplate]    Script Date: 13/02/2019 05:46:23 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spLoadTemplate')
  BEGIN
    DROP PROCEDURE [dbo].[spLoadTemplate]
  END
GO

/****** Object:  StoredProcedure [dbo].[spLoadTemplate]    Script Date: 13/02/2019 05:46:23 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spLoadTemplate]
	@pAnio					INTEGER
	,@pIdTipoPlantilla		INTEGER
	,@pNombre				VARCHAR(150)
	,@pPlantilla			VARBINARY(MAX)
	,@pUsuario				VARCHAR(50)
AS
BEGIN
	DECLARE
		@ContTran			INTEGER
		,@IdError			INTEGER
		,@Message			VARCHAR(500)
		,@Date				DATETIME		=	GETDATE()
		,@Id				INTEGER;

	BEGIN TRY
		IF(@@TRANCOUNT = 0)
		BEGIN
			SET	@ContTran	=	1;
			BEGIN TRAN;
		END;

		UPDATE dbo.tbl_Plantillas
			SET		Activo				=	0
					,Fecha_Modificacion	=	@Date
					,Usuario_Modifico	=	@pUsuario
			WHERE	IdTipoPlantilla		=	@pIdTipoPlantilla
					AND
					Anio				=	@pAnio

		SET	@Id		=	(SELECT ISNULL(MAX(IdPlantilla), 0) + 1 FROM dbo.tbl_Plantillas);

		INSERT INTO dbo.tbl_Plantillas (
				IdPlantilla
				,IdTipoPlantilla
				,Anio
				,Nombre
				,ArchivoPlantilla
				,Activo
				,Fecha_Modificacion
				,Usuario_Modifico
			)
		VALUES (
				@Id
				,@pIdTipoPlantilla
				,@pAnio
				,@pNombre
				,@pPlantilla
				,1
				,@Date
				,@pUsuario
			);

		IF(@ContTran = 1)
			COMMIT TRAN;
	END TRY
	BEGIN CATCH
		IF(@ContTran = 1)
			ROLLBACK TRAN;

		SET	@IdError	=	ERROR_NUMBER();
		SET	@Message	=	CONVERT(VARCHAR(25), @IdError) + ' - ' + ERROR_MESSAGE();

		RAISERROR(@Message, 16, 1);
	END CATCH;
END;
GO

