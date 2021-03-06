/*********************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Algerie Gil
  Stored Procedure: spActualizarComprobaciones
\*********************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spActualizarComprobaciones]    Script Date: 01/03/2019 12:11:25 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spActualizarComprobaciones')
  BEGIN
    DROP PROCEDURE [dbo].[spActualizarComprobaciones]
  END
GO

/****** Object:  StoredProcedure [dbo].[spActualizarComprobaciones]    Script Date: 01/03/2019 12:11:25 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  StoredProcedure [dbo].[spActualizarComprobaciones]    Script Date: 28/2/2019 10:09:18 a.m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO 
CREATE PROCEDURE [dbo].[spActualizarComprobaciones]
(
  @pIdComprobacion      INTEGER
 ,@pIdTipoPlantilla		INTEGER
 ,@pConcepto			VARCHAR(1500)
 ,@pFormula  			VARCHAR(1500)
 ,@pCondicion			VARCHAR(1500)
 ,@pNota     			VARCHAR(500)
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
          INSERT INTO tbl_Comprobaciones (IdComprobacion,IdTipoPlantilla,Concepto,Formula,Condicion,Nota,AdmiteCambios)
          VALUES (@pIdComprobacion,@pIdTipoPlantilla,@pConcepto,@pFormula,@pCondicion,@pNota,1) 
       END
       IF (@pAccion='M') ---Modificar
		 BEGIN         
			 UPDATE tbl_Comprobaciones
			 SET Concepto  = @pConcepto
				,Formula   = @pFormula
				,Condicion = @pCondicion
				,Nota      = @pNota
			 WHERE IdComprobacion=@pIdComprobacion
		  END
     IF (@pAccion='E') ---Eliminar
       BEGIN
         DELETE FROM tbl_Comprobaciones WHERE  IdComprobacion=@pIdComprobacion
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

