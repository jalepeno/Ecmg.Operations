' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  DeleteContentPathsOperation.vb
'  Description :  [type_description_here]
'  Created     :  5/19/2016 9:44:07 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class DeleteContentPathsOperation
  Inherits ActionOperation

#Region "Class Constants"

  Public Const OPERATION_NAME As String = "DeleteContentPaths"
  'Friend Const PARAM_GENERATE_CDF As String = "GenerateCDF"
  'Friend Const PARAM_SAVE_TO_FILE As String = "SaveToFile"
  'Friend Const PARAM_GET_CONTENT As String = "GetContent"
  'Friend Const PARAM_GET_RELATED_DOCUMENTS As String = "GetRelatedDocuments"
  'Friend Const PARAM_METADATA_AS_JSON As String = "SaveMetadataAsJson"
  'Friend Const PARAM_DESTINATION_FOLDER As String = "DestinationFolder"
  'Private Const CTS_DOCS_PATH_REPLACEMENT As String = "%CtsDocsPath%"
  'Private Const BATCH_NAME As String = "%BatchName%"
  'Private Const JOB_NAME As String = "%JobName%"
  'Private Const PROJECT_NAME As String = "%ProjectName%"
  'Private Const PRIMARY_SOURCE_FOLDER_PATH As String = "%PrimarySourceFolderPath%"
  'Private Const DEFAULT_DESTINATION_FOLDER As String = CTS_DOCS_PATH_REPLACEMENT & "\Exports\" & PROJECT_NAME & "\" & JOB_NAME
  ''Private Const DEFAULT_BATCH_DESTINATION_FOLDER As String = CTS_DOCS_PATH_REPLACEMENT & "\Exports\" & BATCH_NAME
  'Friend Const PARAM_CREATE_BATCH_FOLDERS As String = "CreateBatchFolders"
  'Friend Const PARAM_CREATE_ITEM_FOLDERS As String = "CreateItemFolders"
  'Friend Const PARAM_RECREATE_SOURCE_FOLDERS As String = "RecreateSourceFolders"
  'Friend Const PARAM_SAVE_MODE As String = "SaveMode"
  'Private Const DEFAULT_ARCHIVE_PASSWORD As String = ""
  'Friend Const PARAM_ARCHIVE_PASSWORD As String = "ArchivePassword"
  'Friend Const PARAM_GET_ANNOTATIONS As String = "GetAnnotations"
  'Friend Const PARAM_GET_PERMISSIONS As String = "GetPermissions"
  'Friend Const PARAM_VERSION_SCOPE As String = "VersionScope"

#End Region

#Region "Class Variables"

  Private mblnLogResult As Boolean = True
  Private mstrExportedDocumentPath As String

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

#End Region

#Region "Friend Methods"

  Friend Overrides Function OnExecute() As Result
    Try
      Return DeleteContentPaths()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Function DeleteContentPaths() As Result

    Dim lobjECMDocument As Document = Nothing

    Try

      'LogSession.EnterMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))

      RunPreOperationChecks(False)

      lobjECMDocument = Me.WorkItem.Document

      lobjECMDocument.DeleteContentPaths()

      menuResult = OperationEnumerations.Result.Success

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

#End Region

End Class
