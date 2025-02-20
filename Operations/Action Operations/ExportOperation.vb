' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ExportOperation.vb
'  Description :  [type_description_here]
'  Created     :  11/23/2011 8:35:07 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Public Class ExportOperation
  Inherits ActionOperation

#Region "Class Constants"

  Public Const OPERATION_NAME As String = "Export"
  Friend Const PARAM_GENERATE_CDF As String = "GenerateCDF"
  Friend Const PARAM_SAVE_TO_FILE As String = "SaveToFile"
  Friend Const PARAM_GET_CONTENT As String = "GetContent"
  Friend Const PARAM_GET_RELATED_DOCUMENTS As String = "GetRelatedDocuments"
  Friend Const PARAM_METADATA_AS_JSON As String = "SaveMetadataAsJson"
  Friend Const PARAM_DESTINATION_FOLDER As String = "DestinationFolder"
  Private Const CTS_DOCS_PATH_REPLACEMENT As String = "%CtsDocsPath%"
  Private Const BATCH_NAME As String = "%BatchName%"
  Private Const JOB_NAME As String = "%JobName%"
  Private Const PROJECT_NAME As String = "%ProjectName%"
  Private Const PRIMARY_SOURCE_FOLDER_PATH As String = "%PrimarySourceFolderPath%"
  Private Const DEFAULT_DESTINATION_FOLDER As String = CTS_DOCS_PATH_REPLACEMENT & "\Exports\" & PROJECT_NAME & "\" & JOB_NAME
  'Private Const DEFAULT_BATCH_DESTINATION_FOLDER As String = CTS_DOCS_PATH_REPLACEMENT & "\Exports\" & BATCH_NAME
  Friend Const PARAM_CREATE_BATCH_FOLDERS As String = "CreateBatchFolders"
  Friend Const PARAM_CREATE_ITEM_FOLDERS As String = "CreateItemFolders"
  Friend Const PARAM_RECREATE_SOURCE_FOLDERS As String = "RecreateSourceFolders"
  Friend Const PARAM_SAVE_MODE As String = "SaveMode"
  Private Const DEFAULT_ARCHIVE_PASSWORD As String = ""
  Friend Const PARAM_ARCHIVE_PASSWORD As String = "ArchivePassword"
  Friend Const PARAM_GET_ANNOTATIONS As String = "GetAnnotations"
  Friend Const PARAM_GET_PERMISSIONS As String = "GetPermissions"
  Friend Const PARAM_VERSION_SCOPE As String = "VersionScope"

#End Region

#Region "Public Enumerations"

  Public Enum SaveModeEnum
    Archive
    ContentOnly
    MetadataOnly
  End Enum

#End Region

#Region "Class Variables"

  Private ReadOnly mblnLogResult As Boolean = True
  Private mstrExportedDocumentPath As String

#End Region

#Region "Public Properties"

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

  Public ReadOnly Property ExportedDocumentPath As String
    Get
      Return mstrExportedDocumentPath
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

#Region "Friend Methods"

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try
      '' <Added by: Ernie at: 8/13/2014-9:55:16 AM on machine: ERNIE-THINK>
      '' Added temporarily to test work summary output for longer operations.
      'Thread.Sleep(120000)
      '' </Added by: Ernie at: 8/13/2014-9:55:16 AM on machine: ERNIE-THINK>
      Return ExportDocument()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Protected Methods"

  Protected Friend Shared Function GetDefaultSaveParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = New Parameters

      If lobjParameters.Contains(PARAM_DESTINATION_FOLDER) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_DESTINATION_FOLDER,
          DEFAULT_DESTINATION_FOLDER, "Specifies the destination folder to save the file to."))
      End If

      If lobjParameters.Contains(PARAM_SAVE_MODE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_SAVE_MODE, SaveModeEnum.Archive, GetType(SaveModeEnum),
          "Specifies whether or not to save the entire package(Archive), the content only or the metadata only."))
      End If

      If lobjParameters.Contains(PARAM_RECREATE_SOURCE_FOLDERS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_RECREATE_SOURCE_FOLDERS, False,
          "Specifies whether or not to recreate the source folder structure for each document."))
      End If

      If lobjParameters.Contains(PARAM_CREATE_BATCH_FOLDERS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_CREATE_BATCH_FOLDERS, False,
          "Specifies whether or not to create a sub folder for each batch based on the batch name."))
      End If

      If lobjParameters.Contains(PARAM_CREATE_ITEM_FOLDERS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_CREATE_ITEM_FOLDERS, False,
          "Specifies whether or not to create a sub folder for each document based on the document id."))
      End If

      If lobjParameters.Contains(PARAM_METADATA_AS_JSON) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_METADATA_AS_JSON, False,
          "Specifies whether or not the document metadata should be written as json."))
      End If

      If lobjParameters.Contains(PARAM_ARCHIVE_PASSWORD) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_ARCHIVE_PASSWORD, DEFAULT_ARCHIVE_PASSWORD,
          "(Optional) The password to protect the saved package file with.  Only applies when the save mode is set to Archive."))
      End If

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Sub CheckParameters()
    Try
      If Me.Parameters.Contains(PARAM_GENERATE_CDF) = True Then
        Me.Parameters.Remove(Me.Parameters.Item(PARAM_GENERATE_CDF))
      End If
      UpdateParameterToEnum(PARAM_SAVE_MODE, GetType(SaveModeEnum))
      UpdateParameterToEnum(PARAM_VERSION_SCOPE, GetType(VersionScopeEnum))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = New Parameters
      Dim lobjParameter As IParameter = Nothing

      If lobjParameters.Contains(PARAM_SAVE_TO_FILE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_SAVE_TO_FILE, True,
          "Specifies whether or not the exported document should be saved to a file."))
      End If

      If lobjParameters.Contains(PARAM_GET_CONTENT) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_GET_CONTENT, True,
          "Specifies whether or not the document content should be exported. (This parameter is not supported for all providers.)"))
      End If

      If lobjParameters.Contains(PARAM_GET_RELATED_DOCUMENTS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_GET_RELATED_DOCUMENTS, True,
          "Specifies whether or not to look for and retrieve related documents, sometimes known as compound documents. (This parameter is not supported for all providers.)"))
      End If

      ' Add in the save parameters
      For Each lobjParameter In GetDefaultSaveParameters()
        If lobjParameters.Contains(lobjParameter.Name) = False Then
          lobjParameters.Add(lobjParameter)
        End If
      Next

      'If lobjParameters.Contains(PARAM_DESTINATION_FOLDER) = False Then
      '  lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_DESTINATION_FOLDER, _
      '    DEFAULT_DESTINATION_FOLDER, "Specifies the destination folder to save the file to."))
      'End If

      'If lobjParameters.Contains(PARAM_SAVE_MODE) = False Then
      '  lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_SAVE_MODE, SaveMode.Archive, GetType(ExportOperation.SaveMode), _
      '    "Specifies whether or not to save the entire package(Archive), the content only or the metadata only."))
      'End If

      'If lobjParameters.Contains(PARAM_CREATE_BATCH_FOLDERS) = False Then
      '  lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_CREATE_BATCH_FOLDERS, False, _
      '    "Specifies whether or not to create a sub folder for each batch based on the batch name."))
      'End If

      'If lobjParameters.Contains(PARAM_CREATE_ITEM_FOLDERS) = False Then
      '  lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_CREATE_ITEM_FOLDERS, False, _
      '    "Specifies whether or not to create a sub folder for each document based on the document id."))
      'End If

      'If lobjParameters.Contains(PARAM_ARCHIVE_PASSWORD) = False Then
      '  lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_ARCHIVE_PASSWORD, DEFAULT_ARCHIVE_PASSWORD, _
      '    "(Optional) The password to protect the saved package file with.  Only applies when the save mode is set to Archive."))
      'End If

      If lobjParameters.Contains(PARAM_GET_ANNOTATIONS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_GET_ANNOTATIONS, True,
          "Specifies whether or not the document annotations should be exported.  Note that this is only available for some providers.  Those that do not support this feature will ignore this parameter."))
      End If

      If lobjParameters.Contains(PARAM_GET_PERMISSIONS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_GET_PERMISSIONS, True,
          "Specifies whether or not the document permissions should be exported.  Note that this is only available for some providers.  Those that do not support this feature will ignore this parameter."))
      End If

      If lobjParameters.Contains(PARAM_VERSION_SCOPE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_VERSION_SCOPE, VersionScopeEnum.AllVersions,
          GetType(VersionScopeEnum),
          "Specifies which versions of the document should be exported (NOTE: Selective versions are not supported on all export providers."))
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

      ' Look for any property based wildcards and replace them
      SubstitutePropertyWildCards(lstrDestinationFolderPath)

      lstrDestinationFolderPath = Helper.CleanPath(lstrDestinationFolderPath)

      Return lstrDestinationFolderPath

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Sub SubstitutePropertyWildCards(ByRef lpDestinationFolderPath As String)
    Try
      Dim lobjPropertyWildCards As New PropertyWildCards(Me.WorkItem.Document)
      If lobjPropertyWildCards IsNot Nothing Then
        lpDestinationFolderPath = lobjPropertyWildCards.SubstitutePropertyWildCards(lpDestinationFolderPath)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  ''' <summary>
  ''' Exports a single document
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ExportDocument() As Result

    Dim lobjECMDocument As Document = Nothing
    Dim lobjLogResultParameter As IParameter = Me.Parameters.Item(LOG_RESULT)
    Dim lobjExporter As IDocumentExporter
    Dim lstrDestinationFilePath As String = Nothing
    Dim lstrMessage As String

    Try
      'LogSession.EnterMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))

      RunPreOperationChecks(False)
      ' ApplicationLogging.LogInformation($"Exporting document '{WorkItem.SourceDocId}'")

      If String.IsNullOrEmpty(Me.DocumentId) Then
        Throw New DocumentReferenceNotSetException()
      End If

      Dim lobjArgs As New ExportDocumentEventArgs(Me.DocumentId, lobjECMDocument)

      'If Parameters.Contains(PARAM_GENERATE_CDF) AndAlso Parameters.Item(PARAM_GENERATE_CDF).Value = True Then
      '  lobjArgs.GenerateCDF = True
      'Else
      '  lobjArgs.GenerateCDF = False
      'End If

      Dim lenuSaveMode As SaveModeEnum = CType([Enum].Parse(lenuSaveMode.GetType, CStr(Me.Parameters.Item(PARAM_SAVE_MODE).Value)), SaveModeEnum)

      ' lobjArgs.GenerateCDF = GetBooleanParameterValue(PARAM_GENERATE_CDF, False)
      lobjArgs.GetContent = GetBooleanParameterValue(PARAM_GET_CONTENT, True)
      lobjArgs.GetAnnotations = GetBooleanParameterValue(PARAM_GET_ANNOTATIONS, True)
      lobjArgs.GetPermissions = GetBooleanParameterValue(PARAM_GET_PERMISSIONS, True)
      lobjArgs.GetRelatedDocuments = GetBooleanParameterValue(PARAM_GET_RELATED_DOCUMENTS, True)
      'lobjArgs.VersionScope.Scope = CType([Enum].Parse(GetType(VersionScopeEnum), _
      '                                           (GetStringParameterValue(PARAM_VERSION_SCOPE, VersionScopeEnum.AllVersions))), VersionScopeEnum)
      lobjArgs.VersionScope.Scope = GetEnumParameterValue(PARAM_VERSION_SCOPE, GetType(VersionScopeEnum), VersionScopeEnum.AllVersions)
      ' <Modified by: Ernie at 2/14/2014-11:16:32 AM on machine: ERNIE-THINK>
      Select Case Scope
        Case OperationScope.Source
          lobjExporter = CType(SourceConnection.Provider.GetInterface(ProviderClass.DocumentExporter), IDocumentExporter)
        Case OperationScope.Destination
          lobjExporter = CType(DestinationConnection.Provider.GetInterface(ProviderClass.DocumentExporter), IDocumentExporter)
        Case Else
          Throw New InvalidOperationException("Operation scope not set.")
      End Select
      ' </Modified by: Ernie at 2/14/2014-11:16:32 AM on machine: ERNIE-THINK>

      If (lobjExporter IsNot Nothing) Then

        'Transform if we have one
        'If (Me.Batch.Transformations.Count > 0) Then
        '  lobjArgs.Transformation = Me.Batch.Transformations(0)
        'End If

        ' ''This tells the provider to encode the contents but for now it also
        ' ''Tells the provider not to serialize out the document.  In P8 that is the case...
        ' ''Need a new property to tell the provider whether to serialize out or not
        ''lobjArgs.StorageType = Me.Batch.ContentStorageType

        menuResult = ConvertResult(lobjExporter.ExportDocument(lobjArgs))

        If (menuResult = OperationEnumerations.Result.Failed) Then
          If lobjArgs.ErrorMessage.Contains("has zero length") Then
            ProcessedMessage = String.Format("Export Document Failed: {0}", "The requested file is zero length.")
            Throw New DocumentDoesNotExistException(Me.DocumentId, lobjArgs.ErrorMessage, New ZeroLengthContentException(lobjArgs.ErrorMessage))
          Else
            ProcessedMessage = String.Format("Export Document Failed: {0}", lobjArgs.ErrorMessage)
            Throw New DocumentDoesNotExistException(Me.DocumentId, lobjArgs.ErrorMessage)
          End If

        Else

          Me.WorkItem.Document = lobjArgs.Document
          Dim lobjDocument As Document = Me.WorkItem.Document

          If SaveToFile = True Then

            Dim lstrDestinationFolderPath As String = GetCurrentDestinationFolderPath()

            If lstrDestinationFolderPath.Contains(PRIMARY_SOURCE_FOLDER_PATH) Then
              lstrDestinationFolderPath = lobjDocument.CreateAppendedFolderPath(lstrDestinationFolderPath.Replace(PRIMARY_SOURCE_FOLDER_PATH, String.Empty), "\")
            End If

            If IO.Directory.Exists(lstrDestinationFolderPath) = False Then
              IO.Directory.CreateDirectory(lstrDestinationFolderPath)
            End If

            Select Case lenuSaveMode
              Case SaveModeEnum.Archive
                Dim lstrArchivePassword As String = GetStringParameterValue(PARAM_ARCHIVE_PASSWORD, DEFAULT_ARCHIVE_PASSWORD)
                Dim lblnSaveAsJson As Boolean = GetBooleanParameterValue(PARAM_METADATA_AS_JSON, False)
                If lblnSaveAsJson Then
                  lstrDestinationFilePath = String.Format("{0}\{1}.{2}", lstrDestinationFolderPath, lobjDocument.ID, Document.JSON_CONTENT_PACKAGE_FILE_EXTENSION)
                Else
                  lstrDestinationFilePath = String.Format("{0}\{1}.{2}", lstrDestinationFolderPath, lobjDocument.ID, Document.CONTENT_PACKAGE_FILE_EXTENSION)
                End If

                lstrDestinationFilePath = Helper.CleanFile(lstrDestinationFilePath, "~")

                If String.IsNullOrEmpty(lstrArchivePassword) Then
                  Me.WorkItem.Document.Archive(lstrDestinationFilePath)
                Else
                  Me.WorkItem.Document.Archive(lstrDestinationFilePath, True, lstrArchivePassword)
                End If

                Me.ProcessedMessage = String.Format("Document archived to '{0}'.", lstrDestinationFilePath)

              Case SaveModeEnum.ContentOnly

                Dim lintContentCount As Integer = lobjDocument.LatestVersion.ContentCount

                ' Make sure there is content to save
                If lintContentCount = 0 Then
                  Throw New DocumentHasNoContentException(lobjDocument, "Unable to save content to file, document contains no content.")
                End If

                ' Iterate through each content element 
                For Each lobjContent As Content In lobjDocument.LatestVersion.Contents
                  ' Create the destination file path
                  'lstrDestinationFilePath = Helper.CleanPath(String.Format("{0}\{1}", lstrDestinationFolderPath, lobjContent.FileName))

                  ' Temporarily changed for Conagra, they have many files with a content element name of 'file0.pdf'
                  ' Ernie Bahr -- June 29, 2023
                  If lobjContent.FileName = "file0.pdf" Then
                    Dim lobjNameProperty As ECMProperty = Nothing
                    If lobjDocument.LatestVersion.Properties.PropertyExists("DocumentTitle", False, lobjNameProperty) Then
                      Dim lstrCleanFileName As String = Helper.CleanFile(lobjNameProperty.Value, "_", False)
                      lstrDestinationFilePath = Helper.CleanPath(String.Format("{0}\{1}", lstrDestinationFolderPath, lstrCleanFileName))
                      ApplicationLogging.LogInformation(String.Format("Substituted DocumentTitle of '{0}' for content file name of '{1}' for document '{2}'", lstrCleanFileName, lobjContent.FileName, Me.DocumentId))
                    Else
                      lstrDestinationFilePath = Helper.CleanPath(String.Format("{0}\{1}", lstrDestinationFolderPath, lobjContent.FileName))
                    End If
                  Else
                    lstrDestinationFilePath = Helper.CleanPath(String.Format("{0}\{1}", lstrDestinationFolderPath, Helper.CleanFile(lobjContent.FileName, "_", False)))
                  End If


                  Try
                    ' Write the content to a file
                    lobjContent.WriteToFile(lstrDestinationFilePath, True)
                  Catch DirNotFoundEx As System.IO.DirectoryNotFoundException
                    Dim lstrLegalFilePath As String = Helper.ShortenPath(lstrDestinationFilePath)
                    ApplicationLogging.LogWarning(String.Format("Shortened filepath from '{0}' to '{1}' to be able to write to disk.", lstrDestinationFilePath, lstrLegalFilePath))
                    lobjContent.WriteToFile(lstrLegalFilePath, True)
                  End Try
                Next

                Select Case lintContentCount
                  Case 1
                    Me.ProcessedMessage = String.Format("Content file saved to '{0}'.", lstrDestinationFilePath)
                  Case Is > 1
                    Me.ProcessedMessage = String.Format("{0} content files saved to '{1}'.", lintContentCount, lstrDestinationFolderPath)
                End Select

              Case SaveModeEnum.MetadataOnly
                Dim lblnSaveAsJson As Boolean = GetBooleanParameterValue(PARAM_METADATA_AS_JSON, False)
                If lblnSaveAsJson Then
                  lstrDestinationFilePath = String.Format("{0}\{1}.{2}", lstrDestinationFolderPath, lobjDocument.ID, Document.JSON_CONTENT_DEFINITION_FILE_EXTENSION)
                Else
                  lstrDestinationFilePath = String.Format("{0}\{1}.{2}", lstrDestinationFolderPath, lobjDocument.ID, Document.CONTENT_DEFINITION_FILE_EXTENSION)
                End If
                lobjDocument.Save(lstrDestinationFilePath)
                Me.ProcessedMessage = String.Format("Metadata saved to '{0}'.", lstrDestinationFilePath)

            End Select

          End If

        End If

        mstrExportedDocumentPath = lstrDestinationFilePath

        ' Make a notation in the application log
        If String.IsNullOrEmpty(Me.ProcessedMessage) Then
          If Me.SaveToFile = True Then
            lstrMessage = String.Format("Successfully exported document {0}: '{1}' to '{2}'.",
                                                           lobjArgs.Document.ID, lobjArgs.Document.Name, mstrExportedDocumentPath)
            'LogSession.LogMessage(lstrMessage)
            'ApplicationLogging.WriteLogEntry(lstrMessage, TraceEventType.Information, 61204)
          Else
            lstrMessage = String.Format("Successfully exported document {0}: '{1}'.",
                                                           lobjArgs.Document.ID, lobjArgs.Document.Name)
            'LogSession.LogMessage(lstrMessage)
            'ApplicationLogging.WriteLogEntry(lstrMessage, TraceEventType.Information, 61205)
          End If
        End If

        If Me.LogResult = True AndAlso Me.SaveToFile = True Then
          Me.WorkItem.DestinationDocId = mstrExportedDocumentPath
        End If

        'If lobjLogResultParameter IsNot Nothing AndAlso lobjLogResultParameter.Value = True Then
        '  MyBase.EndProcessItem(Migrations.ProcessedStatus.Success, String.Empty, mstrExportedDocumentPath)
        'End If

      Else

        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to get exporter"))
        'If (mblnLogResult) Then
        '  'Unable to get exporter
        '  MyBase.EndProcessItem(Projects.ProcessedStatus.Failed, "Unable to get exporter", String.Empty)
        'End If

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

  Private Shared Function ExportDocument(ByVal lpExporter As IDocumentExporter,
                                  ByVal lpArgs As ExportDocumentEventArgs) As Result

    Try

      Dim lblnReturn As Boolean

      Try
        lblnReturn = lpExporter.ExportDocument(lpArgs)

      Catch ZeroContentEx As ZeroLengthContentException
        ApplicationLogging.WriteLogEntry("Unable to get zero content for document, attempting to export without content instead.", Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 60404)
        lpArgs.GetContent = False
        lblnReturn = lpExporter.ExportDocument(lpArgs)

      Catch LargeContentEx As ContentTooLargeException
        ApplicationLogging.WriteLogEntry("Unable to get large content for document, attempting to export without content instead.", Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 61404)
        lpArgs.GetContent = False
        lblnReturn = lpExporter.ExportDocument(lpArgs)

      Catch ex As Exception
        Throw
      End Try

      If lblnReturn = True Then
        Return OperationEnumerations.Result.Success
      Else
        Return OperationEnumerations.Result.Failed
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Private Sub ExportOperation_Begin(ByVal sender As Object, ByVal e As OperableEventArgs) Handles Me.Begin
    Try
      If Me.WorkItem IsNot Nothing Then
        If Me.Scope = OperationScope.Source Then
          Me.DocumentId = Me.WorkItem.SourceDocId
        Else
          Me.DocumentId = Me.WorkItem.DestinationDocId
        End If
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class