' ********************************************************************************
' '  Document    :  FileOperation.vb
' '  Description :  [type_description_here]
' '  Created     :  11/8/2012-15:19:25
' '  <copyright company="ECMG">
' '      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Public Class FileOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "File"
  Friend Const PARAM_GET_FOLDER_PATHS_FROM_DOCUMENT As String = "GetFolderPathsFromDocument"
  Friend Const PARAM_FOLDER_PATHS As String = "FolderPaths"
  Friend Const PARAM_AUTO_CREATE_FOLDERS As String = "AutoCreateFolders"

#End Region

#Region "Public Overrides Methods"

  Public Overrides ReadOnly Property Name As String
    Get
      Try
        Return OPERATION_NAME
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Overrides ReadOnly Property CanRollback As Boolean
    Get
      Return True
    End Get
  End Property

  Friend Overrides Function OnExecute() As Result
    Try

      Return FileDocument()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Protected Methods"

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = New Parameters

      If lobjParameters.Contains(PARAM_AUTO_CREATE_FOLDERS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_AUTO_CREATE_FOLDERS, True,
          "Specifies whether or not to automatically create the destination folder if it does not exist."))
      End If

      If lobjParameters.Contains(PARAM_GET_FOLDER_PATHS_FROM_DOCUMENT) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_GET_FOLDER_PATHS_FROM_DOCUMENT, False,
          "Specifies whether or not to get the target folder path from the current document."))
      End If

      If lobjParameters.Contains(PARAM_FOLDER_PATHS) = False Then
        'Dim lobjFolderPathsParam As New Parameter
        Dim lobjFolderPathsParam As IParameter = ParameterFactory.Create(PropertyType.ecmString, PARAM_FOLDER_PATHS, Cardinality.ecmMultiValued)
        With lobjFolderPathsParam
          .SystemName = PARAM_FOLDER_PATHS
          .Description = "The list of folder paths in which to file the document."
          .Values.Add("/Sample Folder Path")
        End With
        lobjParameters.Add(lobjFolderPathsParam)
      End If

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Function FileDocument() As Result

    Dim lobjLogResultParameter As IParameter = Me.Parameters.Item(LOG_RESULT)
    Dim lobjFolderManager As IFolderManager = Nothing
    Dim lblnAutoCreateFolders As Boolean
    Dim lblnGetFolderPathsFromDocument As Boolean
    Dim lobjTargetFolderPaths As IList(Of String)

    Try

      RunPreOperationChecks(False)

      If Me.WorkItem IsNot Nothing Then
        If Me.Scope = OperationScope.Source Then
          Me.DocumentId = Me.WorkItem.SourceDocId
          lobjFolderManager = CType(SourceConnection.Provider.GetInterface(ProviderClass.FolderManager), IFolderManager)
        Else
          Me.DocumentId = Me.WorkItem.DestinationDocId
          lobjFolderManager = CType(DestinationConnection.Provider.GetInterface(ProviderClass.FolderManager), IFolderManager)
        End If
      End If

      lblnAutoCreateFolders = CBool(GetParameterValue(PARAM_AUTO_CREATE_FOLDERS, True))
      lblnGetFolderPathsFromDocument = CBool(GetParameterValue(PARAM_GET_FOLDER_PATHS_FROM_DOCUMENT, False))

      If lblnGetFolderPathsFromDocument = True Then
        lobjTargetFolderPaths = Me.WorkItem.Document.FolderPaths
      Else
        Dim lobjFolderPathsParam As IParameter = CType(Parameters, Parameters)(PARAM_FOLDER_PATHS)
        lobjTargetFolderPaths = New List(Of String)
        If lobjFolderPathsParam.HasValue Then
          For Each lobjFolderPath As Object In CType(lobjFolderPathsParam.Values, IEnumerable(Of Object))
            lobjTargetFolderPaths.Add(lobjFolderPath.ToString)
          Next
        Else

        End If
      End If

      If lobjTargetFolderPaths IsNot Nothing AndAlso lobjTargetFolderPaths.Count > 0 Then
        For Each lstrFolderPath As String In lobjTargetFolderPaths
          If lobjFolderManager.FolderPathExists(lstrFolderPath) = False Then
            If lblnAutoCreateFolders Then
              lobjFolderManager.CreateFolder(lstrFolderPath)
            Else
              ProcessedMessage = String.Format("Folder '{0}' does not exist and AutoCreateFolders is set to False.", lstrFolderPath)
              Throw New FolderDoesNotExistException(lstrFolderPath, ProcessedMessage)
            End If
          End If
          lobjFolderManager.FileDocumentByPath(Me.DocumentId, lstrFolderPath)
          menuResult = OperationEnumerations.Result.Success
        Next
        'Else
        '  Beep()
      End If

      ' lobjFolderManager.FolderPathExists()
      ' lobjFolderManager.FileDocumentByPath()
      Return menuResult

    Catch Ex As Exception
      ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, Ex))
      Return menuResult
    End Try
  End Function

#End Region

End Class
