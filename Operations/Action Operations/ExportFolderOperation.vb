'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ExportFolderOperation.vb
'   Description :  [type_description_here]
'   Created     :  3/6/2015 3:46:20 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Public Class ExportFolderOperation
  Inherits ActionOperation

#Region "Class Constants"

  Public Const OPERATION_NAME As String = "ExportFolder"
  Friend Const PARAM_SAVE_TO_FILE As String = "SaveToFile"
  Friend Const PARAM_DESTINATION_FOLDER As String = "DestinationFolder"
  Friend Const PARAM_CREATE_BATCH_FOLDERS As String = "CreateBatchFolders"
  Private Const CTS_DOCS_PATH_REPLACEMENT As String = "%CtsDocsPath%"
  Private Const JOB_NAME As String = "%JobName%"
  Private Const PROJECT_NAME As String = "%ProjectName%"
  Private Const PRIMARY_SOURCE_FOLDER_PATH As String = "%PrimarySourceFolderPath%"
  Private Const DEFAULT_DESTINATION_FOLDER As String = CTS_DOCS_PATH_REPLACEMENT & "\Exports\" & PROJECT_NAME & "\" & JOB_NAME

#End Region

#Region "Constructors"

  Public Sub New()
    Try

      ' Set the default scope
      Scope = OperationScope.Source

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Properties"

  Public Overrides ReadOnly Property CanRollback As Boolean
    Get
      Return False
    End Get
  End Property

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

  Public Property SaveToFile As Boolean
    Get
      Try
        Return Convert.ToBoolean(Parameters.Item(ExportOperation.PARAM_SAVE_TO_FILE).Value)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Boolean)
      Try
        Parameters.Item(ExportOperation.PARAM_SAVE_TO_FILE).Value = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

#End Region

#Region "Friend Methods"

  Friend Overrides Function OnExecute() As Result
    Try
      Return ExportFolder()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Protected Methods"

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = New Parameters
      'Dim lobjParameter As IParameter = Nothing

      If lobjParameters.Contains(PARAM_SAVE_TO_FILE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_SAVE_TO_FILE, True,
          "Specifies whether or not the exported folder should be saved to a file."))
      End If

      If lobjParameters.Contains(PARAM_DESTINATION_FOLDER) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_DESTINATION_FOLDER,
          DEFAULT_DESTINATION_FOLDER, "Specifies the destination folder to save the file to."))
      End If

      If lobjParameters.Contains(PARAM_CREATE_BATCH_FOLDERS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_CREATE_BATCH_FOLDERS, False,
          "Specifies whether or not to create a sub folder for each batch based on the batch name."))
      End If

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Function ExportFolder() As Result

    Dim lobjFolder As Folder
    Dim lobjLogResultParameter As IParameter = Me.Parameters.Item(LOG_RESULT)
    Dim lobjExporter As IFolderExporter
    Dim lstrDestinationFilePath As String = Nothing
    Dim lstrMessage As String

    Try
      'LogSession.EnterMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))

      RunPreOperationChecksForFolder(False)

      Dim lobjArgs As ExportFolderEventArgs

      If Me.DocumentId.Contains("/"c) Then
        lobjArgs = New ExportFolderEventArgs(Me.DocumentId, ExportFolderEventArgs.FolderIdentifierType.FolderPath)
      Else
        lobjArgs = New ExportFolderEventArgs(Me.DocumentId, ExportFolderEventArgs.FolderIdentifierType.FolderId)
      End If

      Select Case Scope
        Case OperationScope.Source
          lobjExporter = CType(SourceConnection.Provider.GetInterface(ProviderClass.FolderExporter), IFolderExporter)
        Case OperationScope.Destination
          lobjExporter = CType(DestinationConnection.Provider.GetInterface(ProviderClass.FolderExporter), IFolderExporter)
        Case Else
          Throw New InvalidOperationException("Operation scope not set.")
      End Select

      If (lobjExporter IsNot Nothing) Then

        menuResult = ConvertResult(lobjExporter.ExportFolder(lobjArgs))

        If (menuResult = OperationEnumerations.Result.Failed) Then
          ProcessedMessage = String.Format("Export Folder Failed: {0}", lobjArgs.ErrorMessage)
          Throw New FolderDoesNotExistException(Me.DocumentId, lobjArgs.ErrorMessage)
        Else

          Me.WorkItem.Folder = lobjArgs.Folder
          lobjFolder = Me.WorkItem.Folder

          If SaveToFile = True Then
            Dim lstrDestinationFolderPath As String = GetCurrentDestinationFolderPath()

            If IO.Directory.Exists(lstrDestinationFolderPath) = False Then
              IO.Directory.CreateDirectory(lstrDestinationFolderPath)
            End If

            lstrDestinationFilePath = String.Format("{0}\{1}.{2}", lstrDestinationFolderPath, lobjFolder.Id, Folder.CONTENT_FOLDER_FILE_EXTENSION)
            lobjFolder.Save(lstrDestinationFilePath)
            Me.ProcessedMessage = String.Format("Folder saved to '{0}'.", lstrDestinationFilePath)

          End If
        End If

        ' Make a notation in the application log
        If Me.SaveToFile = True Then
          lstrMessage = String.Format("Successfully exported folder {0}: '{1}' to '{2}'.",
                                                         lobjArgs.Folder.Id, lobjArgs.Folder.Name, lstrDestinationFilePath)
          'LogSession.LogMessage(lstrMessage)
          'ApplicationLogging.WriteLogEntry(lstrMessage, TraceEventType.Information, 61204)
        Else
          lstrMessage = String.Format("Successfully exported folder {0}: '{1}'.",
                                                         lobjArgs.Folder.Id, lobjArgs.Folder.Name)
          'LogSession.LogMessage(lstrMessage)
          'ApplicationLogging.WriteLogEntry(lstrMessage, TraceEventType.Information, 61205)
        End If

        If Me.LogResult = True AndAlso Me.SaveToFile = True Then
          Me.WorkItem.DestinationDocId = lstrDestinationFilePath
        End If

      Else

        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to get exporter"))

      End If

    Catch ex As Exception
      'LogSession.LogException(ex)
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, ex))
    Finally
      'LogSession.LeaveMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
    End Try

    Return menuResult

  End Function

  Private Function GetCurrentDestinationFolderPath() As String
    Try
      Dim lstrDestinationFolderPath As String = GetStringParameterValue(PARAM_DESTINATION_FOLDER, PARAM_DESTINATION_FOLDER)
      Dim lblnCreateBatchFolders As Boolean = GetBooleanParameterValue(PARAM_CREATE_BATCH_FOLDERS, False)

      lstrDestinationFolderPath = lstrDestinationFolderPath.Replace(CTS_DOCS_PATH_REPLACEMENT, FileHelper.Instance.CtsDocsPath).TrimEnd(CChar("\"))

      If (lstrDestinationFolderPath.Contains(JOB_NAME) OrElse lstrDestinationFolderPath.Contains(PROJECT_NAME)) AndAlso Me.Parent IsNot Nothing AndAlso Me.Parent.GetType.Name = "Batch" Then
        lstrDestinationFolderPath = lstrDestinationFolderPath.Replace(PROJECT_NAME, CType(Me.Parent, Object).Job.Project.Name).Replace(JOB_NAME, CType(Me.Parent, Object).Job.Name).TrimEnd(CChar("\"))
      End If

      If lblnCreateBatchFolders = True Then
        lstrDestinationFolderPath = String.Format(String.Format("{0}\{1}", lstrDestinationFolderPath, Me.Parent.Name))
      End If

      lstrDestinationFolderPath = Helper.CleanPath(lstrDestinationFolderPath)

      Return lstrDestinationFolderPath

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
