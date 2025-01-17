' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  SaveToFileOperation.vb
'  Description :  [type_description_here]
'  Created     :  4/18/2012 1:34:19 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports Documents
Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class SaveToFileOperation
  Inherits ActionOperation

#Region "Class Constants"

  Public Const OPERATION_NAME As String = "SaveToFile"
  Private Const CTS_DOCS_PATH_REPLACEMENT As String = "%CtsDocsPath%"
  Private Const PRIMARY_SOURCE_FOLDER_PATH As String = "%PrimarySourceFolderPath%"
  Private Const DEFAULT_DESTINATION_FOLDER As String = CTS_DOCS_PATH_REPLACEMENT & "\Exports\" & PROJECT_NAME & "\" & JOB_NAME
  Private Const DEFAULT_SAVE_MODE As String = "Archive"
  Private Const DEFAULT_ARCHIVE_PASSWORD As String = ""
  Friend Const PARAM_DESTINATION_FOLDER As String = "DestinationFolder"
  Friend Const PARAM_SAVE_MODE As String = "SaveMode"
  Friend Const PARAM_METADATA_AS_JSON As String = "SaveMetadataAsJson"
  Friend Const PARAM_ARCHIVE_PASSWORD As String = "ArchivePassword"
  Friend Const PARAM_CREATE_ITEM_FOLDERS As String = "CreateItemFolders"
  Friend Const PARAM_CREATE_BATCH_FOLDERS As String = "CreateBatchFolders"
  Friend Const PARAM_RECREATE_SOURCE_FOLDERS As String = "RecreateSourceFolders"
  Friend Const PARAM_CLEAR_CONTENT_PATH As String = "ClearContentPath"
  Private Const JOB_NAME As String = "%JobName%"
  Private Const PROJECT_NAME As String = "%ProjectName%"

#End Region

#Region "Public Enumerations"

  Public Enum SaveMode
    Archive
    ContentOnly
    MetadataOnly
  End Enum

#End Region

#Region "Class Variables"

  Private mstrDestinationFolderPath As String = String.Empty

#End Region

#Region "Public Properties"

  Public Overrides ReadOnly Property Name As String
    Get
      Return OPERATION_NAME
    End Get
  End Property

  Public Overrides ReadOnly Property CanRollback As Boolean
    Get
      Return False
    End Get
  End Property

  Public ReadOnly Property CurrentDestinationFolderPath As String
    Get
      Try
        Return GetCurrentDestinationFolderPath()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Property DestinationFolderPath As String
    Get
      Try
        If String.IsNullOrEmpty(mstrDestinationFolderPath) Then
          mstrDestinationFolderPath = GetStringParameterValue(PARAM_DESTINATION_FOLDER, DEFAULT_DESTINATION_FOLDER)
        End If
        Return mstrDestinationFolderPath
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As String)
      Try
        mstrDestinationFolderPath = value
        Parameters.Item(PARAM_DESTINATION_FOLDER).Value = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

#End Region

#Region "Friend Methods"

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try

      If Not String.IsNullOrEmpty(SaveToFile()) Then
        Return Result.Success
      Else
        If String.IsNullOrEmpty(Me.ProcessedMessage) Then
          Me.ProcessedMessage = "Failed to save to file"
        End If
        Return Result.Failed
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Protected Methods"

  Public Overrides Sub CheckParameters()
    Try
      UpdateParameterToEnum(PARAM_SAVE_MODE, GetType(ExportOperation.SaveModeEnum))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = New Parameters

      If lobjParameters.Contains(PARAM_DESTINATION_FOLDER) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_DESTINATION_FOLDER,
          DEFAULT_DESTINATION_FOLDER, "Specifies the destination folder to save the file to."))
      End If

      ' Add in the save parameters
      For Each lobjParameter In ExportOperation.GetDefaultSaveParameters()
        If lobjParameters.Contains(lobjParameter.Name) = False Then
          lobjParameters.Add(lobjParameter)
        End If
      Next

      If lobjParameters.Contains(PARAM_CLEAR_CONTENT_PATH) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_CLEAR_CONTENT_PATH,
          False, "Specifies whether or not to clear the underlying content path when complete.  This should be false except in special cases such as rendering large files."))
      End If

      'If lobjParameters.Contains(PARAM_SAVE_MODE) = False Then
      '  lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_SAVE_MODE, SaveMode.Archive, GetType(SaveMode), _
      '    "Specifies whether or not to save the entire package(Archive), the content only or the metadata only."))
      'End If

      'If lobjParameters.Contains(PARAM_CREATE_ITEM_FOLDERS) = False Then
      '  lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_CREATE_ITEM_FOLDERS, False, _
      '    "Specifies whether or not to create a sub folder for each document based on the document id."))
      'End If

      'If lobjParameters.Contains(PARAM_ARCHIVE_PASSWORD) = False Then
      '  lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_ARCHIVE_PASSWORD, DEFAULT_ARCHIVE_PASSWORD, _
      '    "(Optional) The password to protect the saved package file with.  Only applies when the save mode is set to Archive."))
      'End If

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Overridable Function SaveToFile() As String

    Dim lstrDestinationFilePath As String = Nothing

    Try

      If Me.WorkItem Is Nothing Then
        Throw New InvalidOperationException("The work item is not initialized.")
      End If

      If Me.WorkItem.Document Is Nothing Then
        Throw New Exceptions.DocumentReferenceNotSetException
      End If

      Dim lenuSaveMode As SaveMode = CType([Enum].Parse(lenuSaveMode.GetType, CStr(Me.Parameters.Item(PARAM_SAVE_MODE).Value)), SaveMode)
      Dim lblnCreateItemFolders As Boolean = GetBooleanParameterValue(PARAM_CREATE_ITEM_FOLDERS, False)
      Dim lblnClearContentPath As Boolean = GetBooleanParameterValue(PARAM_CLEAR_CONTENT_PATH, False)

      'Dim lstrDestinationFolderPath As String = GetStringParameterValue(PARAM_DESTINATION_FOLDER, PARAM_DESTINATION_FOLDER)
      'lstrDestinationFolderPath = lstrDestinationFolderPath.Replace(CTS_DOCS_PATH_REPLACEMENT, FileHelper.Instance.CtsDocsPath).TrimEnd(CChar("\"))

      Dim lobjDocument As Document = Me.WorkItem.Document

      'If lblnCreateItemFolders = True Then
      '  lstrDestinationFolderPath = String.Format("{0}\{1}", lstrDestinationFolderPath, lobjDocument.ID)
      'End If

      'If Directory.Exists(lstrDestinationFolderPath) = False Then
      '  Directory.CreateDirectory(lstrDestinationFolderPath)
      'End If

      Dim lstrCurrentDestinationFolderPath As String = GetCurrentDestinationFolderPath()

      If lstrCurrentDestinationFolderPath.Contains(PRIMARY_SOURCE_FOLDER_PATH) Then
        lstrCurrentDestinationFolderPath = lobjDocument.CreateAppendedFolderPath(lstrCurrentDestinationFolderPath.Replace(PRIMARY_SOURCE_FOLDER_PATH, String.Empty), "\")
      End If

      If Directory.Exists(lstrCurrentDestinationFolderPath) = False Then
        Directory.CreateDirectory(lstrCurrentDestinationFolderPath)
      End If

      Select Case lenuSaveMode
        Case SaveMode.Archive
          Dim lstrArchivePassword As String = GetStringParameterValue(PARAM_ARCHIVE_PASSWORD, DEFAULT_ARCHIVE_PASSWORD)
          Dim lblnSaveAsJson As Boolean = GetBooleanParameterValue(PARAM_METADATA_AS_JSON, False)
          If lblnSaveAsJson Then
            lstrDestinationFilePath = String.Format("{0}\{1}.{2}", lstrCurrentDestinationFolderPath, lobjDocument.ID, Document.JSON_CONTENT_PACKAGE_FILE_EXTENSION)
          Else
            lstrDestinationFilePath = String.Format("{0}\{1}.{2}", lstrCurrentDestinationFolderPath, lobjDocument.ID, Document.CONTENT_PACKAGE_FILE_EXTENSION)
          End If
          If String.IsNullOrEmpty(lstrArchivePassword) Then
            Me.WorkItem.Document.Archive(lstrDestinationFilePath)
          Else
            Me.WorkItem.Document.Archive(lstrDestinationFilePath, True, lstrArchivePassword)
          End If

          Dim lstrMessage As String = $"Document archived to '{lstrDestinationFilePath}'."
          If String.IsNullOrEmpty(Me.ProcessedMessage) Then
            Me.ProcessedMessage = lstrMessage
          Else
            Me.ProcessedMessage = $"{Me.ProcessedMessage}, {lstrMessage}"
          End If

          If lblnClearContentPath = True Then
            ' Try to clear the content path(s) if applicable.
            For Each lobjVersion As Version In Me.WorkItem.Document.Versions
              For Each lobjContent As Content In lobjVersion.Contents
                Try
                  If lobjContent.Data.StreamType = "FileStream" AndAlso File.Exists(lobjContent.ContentPath) Then
                    File.Delete(lobjContent.ContentPath)
                  End If
                Catch ex As Exception
                  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
                  Continue For
                End Try
              Next
            Next
          End If
        Case SaveMode.ContentOnly

          Dim lintContentCount As Integer = lobjDocument.LatestVersion.ContentCount

          ' Make sure there is content to save
          If lintContentCount = 0 Then
            Throw New Exceptions.DocumentHasNoContentException(lobjDocument, "Unable to save content to file, document contains no content.")
          End If

          ' Iterate through each content element 
          For Each lobjContent As Content In lobjDocument.LatestVersion.Contents
            ' Create the destination file path
            lstrDestinationFilePath = Helper.CleanPath(String.Format("{0}\{1}",
                                        lstrCurrentDestinationFolderPath, lobjContent.FileName))

            ' Write the content to a file
            lobjContent.WriteToFile(lstrDestinationFilePath, True)
          Next

          Select Case lintContentCount
            Case 1
              Me.ProcessedMessage = String.Format("Content file saved to '{0}'.", lstrDestinationFilePath)
            Case Is > 1
              Me.ProcessedMessage = String.Format("{0} content files saved to '{1}'.", lintContentCount, lstrCurrentDestinationFolderPath)
          End Select

        Case SaveMode.MetadataOnly
          Dim lblnSaveAsJson As Boolean = GetBooleanParameterValue(PARAM_METADATA_AS_JSON, False)
          If lblnSaveAsJson Then
            lstrDestinationFilePath = String.Format("{0}\{1}.{2}", lstrCurrentDestinationFolderPath, lobjDocument.ID, Document.JSON_CONTENT_DEFINITION_FILE_EXTENSION)
          Else
            lstrDestinationFilePath = String.Format("{0}\{1}.{2}", lstrCurrentDestinationFolderPath, lobjDocument.ID, Document.CONTENT_DEFINITION_FILE_EXTENSION)
          End If

          lobjDocument.Save(lstrDestinationFilePath)
          Me.ProcessedMessage = String.Format("Metadata saved to '{0}'.", lstrDestinationFilePath)
      End Select

      If ((String.IsNullOrEmpty(Me.WorkItem.DestinationDocId)) AndAlso (Not String.IsNullOrEmpty(lstrDestinationFilePath))) Then
        Me.WorkItem.DestinationDocId = lstrDestinationFilePath
      End If

      Return lstrDestinationFilePath

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, ex))
      Return String.Empty
    End Try

  End Function

  Private Function GetCurrentDestinationFolderPath() As String
    Try
      Dim lstrDestinationFolderPath As String
      Dim lblnCreateBatchFolders As Boolean = GetBooleanParameterValue(PARAM_CREATE_BATCH_FOLDERS, False)
      Dim lblnCreateItemFolders As Boolean = GetBooleanParameterValue(PARAM_CREATE_ITEM_FOLDERS, False)
      Dim lblnRecreateSourceFolders As Boolean = GetBooleanParameterValue(PARAM_RECREATE_SOURCE_FOLDERS, False)

      lstrDestinationFolderPath = GetStringParameterValue(PARAM_DESTINATION_FOLDER, PARAM_DESTINATION_FOLDER)
      lstrDestinationFolderPath = lstrDestinationFolderPath.Replace(CTS_DOCS_PATH_REPLACEMENT, FileHelper.Instance.CtsDocsPath).TrimEnd(CChar("\"))

      If (lstrDestinationFolderPath.Contains(JOB_NAME) OrElse lstrDestinationFolderPath.Contains(PROJECT_NAME)) AndAlso Me.Parent IsNot Nothing AndAlso Me.Parent.GetType.Name = "Batch" Then
        lstrDestinationFolderPath = lstrDestinationFolderPath.Replace(PROJECT_NAME, CType(Me.Parent, Object).Job.Project.Name).Replace(JOB_NAME, CType(Me.Parent, Object).Job.Name).TrimEnd(CChar("\"))
      End If

      If lblnRecreateSourceFolders = True Then
        If lstrDestinationFolderPath.EndsWith("\"c) Then
          lstrDestinationFolderPath = String.Format("{0}{1}", lstrDestinationFolderPath, PRIMARY_SOURCE_FOLDER_PATH)
        Else
          lstrDestinationFolderPath = String.Format("{0}\{1}", lstrDestinationFolderPath, PRIMARY_SOURCE_FOLDER_PATH)
        End If
      Else
        If lblnCreateBatchFolders = True Then
          lstrDestinationFolderPath = String.Format(String.Format("{0}\{1}", lstrDestinationFolderPath, Me.Parent.Name))
        End If

        If lblnCreateItemFolders = True Then
          lstrDestinationFolderPath = String.Format("{0}\{1}", lstrDestinationFolderPath, Me.DocumentId)
        End If
      End If

      lstrDestinationFolderPath = Helper.CleanPath(lstrDestinationFolderPath)

      Return lstrDestinationFolderPath

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
