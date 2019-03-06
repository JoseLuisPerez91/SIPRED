/*****************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Algerie Gil
  Stored Procedure: spActualizarCruces
\*****************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spObtenerCruces]    Script Date: 25/02/2019 06:04:43 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spActualizarCruces')
  BEGIN
    DROP PROCEDURE [dbo].[spActualizarCruces]
  END
GO

/****** Object:  StoredProcedure [dbo].[spActualizarCruces]    Script Date: 25/02/2019 06:04:43 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[spActualizarCruces]
(
  @pIdCruce		        INTEGER
 ,@pIdTipoPlantilla		INTEGER
 ,@pConcepto			VARCHAR(1500)
 ,@pFormula  			VARCHAR(1500)
 ,@pCondicion			VARCHAR(1500)
 ,@pNota     			VARCHAR(500)
 ,@pLecImportes         INTEGER
 ,@pAccion              VARCHAR(1)
)
AS
 BEGIN
   DECLARE @ContTran AS INTEGER
   DECLARE @IdError	 AS INTEGER
   DECLARE @Message	 AS VARCHAR(500)
   DECLARE @Date     AS DATETIME = GETDATE()
   DECLARE @Id       AS INTEGER;

   BEGIN TRY
     IF(@@TRANCOUNT = 0)
       BEGIN
         SET @ContTran = 1;
         BEGIN TRAN;
       END;
     IF(@pAccion='I') ---Insertar
       BEGIN
         INSERT INTO tbl_Cruces(IdCruce, IdTipoPlantilla, Concepto, Formula, Condicion, Nota, LecturaImportes)
         VALUES(@pIdCruce, @pIdTipoPlantilla, @pConcepto, @pFormula, @pCondicion, @pNota, @pLecImportes)  
       END
     IF (@pAccion='M') ---Modificar
       BEGIN
         UPDATE tbl_Cruces
         SET Concepto = @pConcepto
             ,Formula = @pFormula
             ,Condicion = @pCondicion
             ,Nota = @pNota
             ,LecturaImportes = @pLecImportes
         WHERE IdCruce=@pIdCruce
       END
     IF (@pAccion='E') ---Eliminar
       BEGIN
         DELETE FROM tbl_Cruces WHERE IdCruce=@pIdCruce
       END
     IF (@pAccion<>'')
	   BEGIN
         UPDATE tbl_Plantillas SET Fecha_Modificacion = GETDATE() WHERE IdPlantilla = @pIdTipoPlantilla AND Activo = 1
       END
     IF(@ContTran = 1)
       BEGIN
         COMMIT TRAN;
       END
   END TRY
   BEGIN CATCH
     IF(@ContTran = 1)
	   BEGIN
         ROLLBACK TRAN;
       END
     SET @IdError = ERROR_NUMBER();
     SET @Message = CONVERT(VARCHAR(25), @IdError) + ' - ' + ERROR_MESSAGE();

     RAISERROR(@Message, 16, 1);
	END CATCH;
 END;
GO


