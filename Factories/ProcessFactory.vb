' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ProcessFactory.vb
'  Description :  [type_description_here]
'  Created     :  11/29/2011 4:30:32 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Transformations
Imports Documents.Utilities

#End Region

Public Class ProcessFactory

#Region "Constructors"

  Private Sub New()

  End Sub

#End Region

#Region "Public Methods"

  Public Shared Function CreateEmptyProcess(ByVal lpName As String) As IProcess
    Try
      Return New Process(lpName, lpName, String.Empty)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function CreateEmptyProcess(ByVal lpName As String, ByVal lpDisplayName As String, ByVal lpDescription As String) As IProcess
    Try
      Return New Process(lpName, lpDisplayName, lpDescription)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  ''' Creates a single operation process based on the operation type.
  ''' </summary>
  ''' <param name="lpOperationType">The operation type</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function CreateFromOperation(ByVal lpOperationType As String) As IProcess
    Try

      If String.IsNullOrEmpty(lpOperationType) Then
        Throw New ArgumentNullException("lpOperation")
      End If

      Dim lobjOperation As IOperable = OperationFactory.Create(lpOperationType)

      If TypeOf lobjOperation Is IOperation Then
        Return CreateFromOperation(CType(lobjOperation, IOperation))
      ElseIf TypeOf lobjOperation Is IProcess Then
        Return CType(lobjOperation, IProcess)
      Else
        Throw New OperationException(lobjOperation, String.Format("Failed to create '{0}' operation.", lpOperationType))
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  ''' Creates a new process object from a single operation
  ''' </summary>
  ''' <param name="lpOperation">The operation</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function CreateFromOperation(ByVal lpOperation As IOperation) As IProcess
    Try

      If lpOperation Is Nothing Then
        Throw New ArgumentNullException("lpOperation")
      End If

      Dim lstrProcessName As String = String.Format("Single {0} Operation Process", lpOperation.Name)
      Dim lobjProcess As New Process(lstrProcessName, lstrProcessName, String.Empty)

      lobjProcess.Operations.Add(lpOperation)

      Return lobjProcess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function CreateMigrationProcess(ByVal lpTransformation As Transformation) As IProcess
    Try

      ' Create a migration process consisting of an export and import and optionally a transformation
      Dim lstrProcessName As String = "Migration Process"
      Dim lobjExportOperation As ExportOperation = Nothing
      Dim lobjTransformOperation As TransformOperation = Nothing
      Dim lobjImportOperation As IOperation = Nothing

      Dim lobjMigrateProcess As IProcess

      lobjExportOperation = CType(OperationFactory.Create("Export"), ExportOperation)
      lobjExportOperation.Parameters.Item(ExportOperation.PARAM_SAVE_TO_FILE).Value = False

      lobjImportOperation = CType(OperationFactory.Create("Import"), IOperation)
      lobjImportOperation.Scope = OperationScope.Destination

      If lpTransformation IsNot Nothing Then
        lobjTransformOperation = CType(OperationFactory.Create("Transform"), TransformOperation)
        lobjTransformOperation.Transformations.Add(lpTransformation)
      End If

      ' Create the process starting with the export operation
      lobjMigrateProcess = New Process(lstrProcessName, lstrProcessName, String.Empty)

      ' Add the export operation
      lobjMigrateProcess.Operations.Add(lobjExportOperation)

      ' If a transformation was specified, add the transformation operation
      If lobjTransformOperation IsNot Nothing Then
        lobjMigrateProcess.Operations.Add(lobjTransformOperation)
      End If

      ' Add the import operation
      lobjMigrateProcess.Operations.Add(lobjImportOperation)

      'lobjMigrateProcess.Name = "Migration Process"

      Return lobjMigrateProcess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function CreateMigrateCustomObjectProcess(ByVal lpTransformation As Transformation) As IProcess
    Try

      ' Create a migration process consisting of an export and import and optionally a transformation
      Dim lstrProcessName As String = "Migrate CustomObject Process"
      Dim lobjExportOperation As ExportCustomObjectOperation = Nothing
      'Dim lobjTransformOperation As TransformFolderOperation = Nothing
      Dim lobjImportOperation As IOperation = Nothing

      Dim lobjMigrateProcess As IProcess

      lobjExportOperation = CType(OperationFactory.Create("ExportCustomObject"), ExportCustomObjectOperation)
      lobjExportOperation.Parameters.Item(ExportOperation.PARAM_SAVE_TO_FILE).Value = False

      lobjImportOperation = CType(OperationFactory.Create("ImportCustomObject"), IOperation)
      lobjImportOperation.Scope = OperationScope.Destination

      'If lpTransformation IsNot Nothing Then
      '  lobjTransformOperation = CType(OperationFactory.Create("TransformFolder"), TransformFolderOperation)
      '  lobjTransformOperation.Transformations.Add(lpTransformation)
      'End If

      ' Create the process starting with the export operation
      lobjMigrateProcess = New Process(lstrProcessName, lstrProcessName, String.Empty)

      ' Add the export operation
      lobjMigrateProcess.Operations.Add(lobjExportOperation)

      ' If a transformation was specified, add the transformation operation
      'If lobjTransformOperation IsNot Nothing Then
      '  lobjMigrateProcess.Operations.Add(lobjTransformOperation)
      'End If

      ' Add the import operation
      lobjMigrateProcess.Operations.Add(lobjImportOperation)

      'lobjMigrateProcess.Name = "Migration Process"

      Return lobjMigrateProcess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function CreateMigrateFolderProcess(ByVal lpTransformation As Transformation) As IProcess
    Try

      ' Create a migration process consisting of an export and import and optionally a transformation
      Dim lstrProcessName As String = "Migrate Folder Process"
      Dim lobjExportOperation As ExportFolderOperation = Nothing
      Dim lobjTransformOperation As TransformFolderOperation = Nothing
      Dim lobjImportOperation As IOperation = Nothing

      Dim lobjMigrateProcess As IProcess

      lobjExportOperation = CType(OperationFactory.Create("ExportFolder"), ExportFolderOperation)
      lobjExportOperation.Parameters.Item(ExportOperation.PARAM_SAVE_TO_FILE).Value = False

      lobjImportOperation = CType(OperationFactory.Create("ImportFolder"), IOperation)
      lobjImportOperation.Scope = OperationScope.Destination

      If lpTransformation IsNot Nothing Then
        lobjTransformOperation = CType(OperationFactory.Create("TransformFolder"), TransformFolderOperation)
        lobjTransformOperation.Transformations.Add(lpTransformation)
      End If

      ' Create the process starting with the export operation
      lobjMigrateProcess = New Process(lstrProcessName, lstrProcessName, String.Empty)

      ' Add the export operation
      lobjMigrateProcess.Operations.Add(lobjExportOperation)

      ' If a transformation was specified, add the transformation operation
      If lobjTransformOperation IsNot Nothing Then
        lobjMigrateProcess.Operations.Add(lobjTransformOperation)
      End If

      ' Add the import operation
      lobjMigrateProcess.Operations.Add(lobjImportOperation)

      'lobjMigrateProcess.Name = "Migration Process"

      Return lobjMigrateProcess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class