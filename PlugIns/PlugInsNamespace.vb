'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Arguments

#End Region

Namespace PlugIns


#Region "Event Delegates"
  ''' <summary>
  ''' Raised from PlugIn when beginning to Execute
  ''' </summary>
  ''' <param name="lpMessage"></param>
  ''' <remarks></remarks>
  Public Delegate Sub ExecuteBeginHandler(ByVal lpMessage As String)

  ''' <summary>
  ''' Raised from PlugIn to report progress of Execution
  ''' </summary>
  ''' <param name="lpPercentProgress"></param>
  ''' <param name="lpMessage"></param>
  ''' <remarks></remarks>
  Public Delegate Sub ExecuteReportProgressHandler(ByVal lpPercentProgress As Integer, ByVal lpMessage As String)

  ''' <summary>
  ''' Raised from PlugIn to signal completion of Execute
  ''' </summary>
  ''' <param name="lpMessage"></param>
  ''' <remarks></remarks>
  Public Delegate Sub ExecuteCompleteHandler(ByVal lpMessage As String)

  ''' <summary>
  ''' Raised from PlugIn to signal Execute Error
  ''' </summary>
  ''' <param name="lpMessage"></param>
  ''' <remarks></remarks>
  Public Delegate Sub ExecuteErrorHandler(ByVal lpMessage As String)

  ' For IDocumentPlugIn
  ''' <summary>Delegate event handler for the BeginProcess event.</summary>
  Public Delegate Function BeginProcessEventHandler(ByVal sender As Object, ByRef e As DocumentEventArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the EndProcess event.</summary>
  Public Delegate Function EndProcessEventHandler(ByVal sender As Object, ByRef e As DocumentEventArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the BeginExportDocument event.</summary>
  Public Delegate Function BeginExportDocumentEventHandler(ByVal sender As Object, ByRef e As ExportDocumentEventArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the ExportDocumentError event.</summary>
  Public Delegate Function ExportDocumentErrorEventHandler(ByVal sender As Object, ByRef e As DocumentExportErrorEventArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the EndExportDocument event.</summary>
  Public Delegate Function EndExportDocumentEventHandler(ByVal sender As Object, ByRef e As ExportDocumentEventArgs) As PlugInExecuteReturnArgs

  ' <Removed by: Ernie at: 9/29/2014-2:25:10 PM on machine: ERNIE-THINK>
  '   ''' <summary>Delegate event handler for the BeginExportDocuments event.</summary>
  '   Public Delegate Function BeginExportDocumentsEventHandler(ByVal sender As Object, ByRef e As ExportDocumentsEventArgs) As PlugInExecuteReturnArgs
  ' 
  '   ''' <summary>Delegate event handler for the EndExportDocuments event.</summary>
  '   Public Delegate Function EndExportDocumentsEventHandler(ByVal sender As Object, ByRef e As ExportDocumentsEventArgs) As PlugInExecuteReturnArgs
  ' 
  '   ''' <summary>Delegate event handler for the BeginExportFolder event.</summary>
  '   Public Delegate Function BeginExportFolderEventHandler(ByVal sender As Object, ByRef e As ExportFolderEventArgs) As PlugInExecuteReturnArgs
  ' 
  '   ''' <summary>Delegate event handler for the EndExportFolder event.</summary>
  '   Public Delegate Function EndExportFolderEventHandler(ByVal sender As Object, ByRef e As ExportFolderEventArgs) As PlugInExecuteReturnArgs
  ' </Removed by: Ernie at: 9/29/2014-2:25:10 PM on machine: ERNIE-THINK>

  ''' <summary>Delegate event handler for the BeginTransform event.</summary>
  Public Delegate Function BeginTransformDocumentEventHandler(ByVal sender As Object, ByRef e As TransformDocumentEventArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the EndTransform event.</summary>
  Public Delegate Function EndTransformDocumentEventHandler(ByVal sender As Object, ByRef e As TransformDocumentEventArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the BeginImportDocument event.</summary>
  Public Delegate Function BeginImportDocumentEventHandler(ByVal sender As Object, ByRef e As ImportDocumentArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the ImportDocumentError event.</summary>
  Public Delegate Function ImportDocumentErrorEventHandler(ByVal sender As Object, ByRef e As DocumentImportErrorEventArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the EndImportDocument event.</summary>
  Public Delegate Function EndImportDocumentEventHandler(ByVal sender As Object, ByRef e As ImportDocumentArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the BeginImportDocument event.</summary>
  Public Delegate Function BeginImportRecordEventHandler(ByVal sender As Object, ByRef e As ImportDocumentArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the ImportRecordError event.</summary>
  Public Delegate Function ImportRecordErrorEventHandler(ByVal sender As Object, ByRef e As RecordImportErrorEventArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the EndImportRecord event.</summary>
  Public Delegate Function EndImportRecordEventHandler(ByVal sender As Object, ByRef e As ImportDocumentArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the BeginMigrateDocument event.</summary>
  Public Delegate Function BeginMigrateDocumentEventHandler(ByVal sender As Object, ByRef e As MigrateDocumentEventArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the MigrateDocumentError event.</summary>
  Public Delegate Function MigrateDocumentErrorEventHandler(ByVal sender As Object, ByRef e As MigrateDocumentErrorEventArgs) As PlugInExecuteReturnArgs

  ''' <summary>Delegate event handler for the EndMigrateDocument event.</summary>
  Public Delegate Function EndMigrateDocumentEventHandler(ByVal sender As Object, ByRef e As MigrateDocumentEventArgs) As PlugInExecuteReturnArgs

#End Region

#Region "Enums"
  ''' <summary>
  ''' Sync or Async
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum PlugInExecutionMode
    Synchronous = 0
    Asynchronous = 1
  End Enum

  ''' <summary>
  ''' Tells caller when to execute the plugin (i.e. when to call Execute() method)
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum PlugInExecuteTiming
    BeforeDocumentAdd = 0
    AfterDocumentAdd = 1
    BothBeforeAndAfterAdd = 2
    AnyTime = 3
  End Enum

  ''' <summary>
  ''' Tells the caller what the plug in does (at a high level)
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum PlugInType
    GeneratesCDF = 0
    PreProcessor = 1
    PostProcessor = 2
  End Enum

#End Region


  ''' <summary>
  ''' IPlugIn Interface
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IPlugIn

#Region "Properties"
    ReadOnly Property Name() As String
    ReadOnly Property Version() As String
    ReadOnly Property Description() As String
    ReadOnly Property PlugInClassName() As String
    ReadOnly Property PlugInExecutionMode() As PlugInExecutionMode
    ReadOnly Property PlugInExecuteTiming() As PlugInExecuteTiming
    ReadOnly Property PlugInPropertyDefinitions() As PlugInPropertyDefinitions
    Property PlugInProperties() As PlugInProperties
    ReadOnly Property PlugInType() As PlugInType
#End Region

#Region "Methods"
    ''' <summary>
    ''' Called to execute the main functionality of the plugin
    ''' </summary>
    ''' <remarks></remarks>
    Function Execute() As PlugInExecuteReturnArgs

#End Region

#Region "Events"
    Event ExecuteBegin As ExecuteBeginHandler
    Event ExecuteReportProgress As ExecuteReportProgressHandler
    Event ExecuteComplete As ExecuteCompleteHandler
    Event ExecuteError As ExecuteErrorHandler
#End Region

  End Interface

  Public Interface IDocumentPlugin

#Region "Methods"

    '' Overall Document Processing
    'Function BeginProcess() As BeginProcessEventHandler
    'Function EndProcess() As EndProcessEventHandler

    '' Export Document
    'Function BeginExportDocument() As BeginExportDocumentEventHandler
    'Function ExportDocumentError() As ExportDocumentErrorEventHandler
    'Function EndExportDocument() As EndExportDocumentEventHandler

    '' Export Documents
    'Function BeginExportDocuments() As BeginExportDocumentsEventHandler
    'Function EndExportDocuments() As EndExportDocumentsEventHandler

    '' Export Folder
    'Function BeginExportFolder() As BeginExportFolderEventHandler
    'Function EndExportFolder() As EndExportFolderEventHandler

    '' Transform Document
    'Function BeginTransformDocument() As BeginTransformDocumentEventHandler
    'Function EndTransformDocument() As EndTransformDocumentEventHandler

    '' Import Document
    'Function BeginImportDocument() As BeginImportDocumentEventHandler
    'Function ImportDocumentError() As ImportDocumentErrorEventHandler
    'Function EndImportDocument() As EndImportDocumentEventHandler

    '' Import Record
    'Function BeginImportRecord() As BeginImportRecordEventHandler
    'Function ImportRecordError() As ImportRecordErrorEventHandler
    'Function EndImportRecord() As EndImportRecordEventHandler

    '' Migrate Document
    'Function BeginMigrateDocument() As BeginMigrateDocumentEventHandler
    'Function MigrateDocumentError() As MigrateDocumentErrorEventHandler
    'Function EndMigrateDocument() As EndMigrateDocumentEventHandler

    '' Overall Document Processing
    ''' <summary>Event handler for the BeginProcess event.</summary>
    Function BeginProcessEventHandler(ByVal sender As Object, ByRef e As DocumentEventArgs) As PlugInExecuteReturnArgs
    ''' <summary>Event handler for the EndProcess event.</summary>
    Function EndProcessEventHandler(ByVal sender As Object, ByRef e As DocumentEventArgs) As PlugInExecuteReturnArgs

    '' Export Document
    ''' <summary>Event handler for the BeginExportDocument event.</summary>
    Function BeginExportDocumentEventHandler(ByVal sender As Object, ByRef e As ExportDocumentEventArgs) As PlugInExecuteReturnArgs
    ''' <summary>Event handler for the ExportDocumentError event.</summary>
    Function ExportDocumentErrorEventHandler(ByVal sender As Object, ByRef e As DocumentExportErrorEventArgs) As PlugInExecuteReturnArgs
    ''' <summary>Event handler for the EndExportDocument event.</summary>
    Function EndExportDocumentEventHandler(ByVal sender As Object, ByRef e As ExportDocumentEventArgs) As PlugInExecuteReturnArgs

    ' <Removed by: Ernie at: 9/29/2014-2:24:19 PM on machine: ERNIE-THINK>
    '     '' Export Documents
    '     ''' <summary>Event handler for the BeginExportDocuments event.</summary>
    '     Function BeginExportDocumentsEventHandler(ByVal sender As Object, ByRef e As ExportDocumentsEventArgs) As PlugInExecuteReturnArgs
    '     ''' <summary>Event handler for the EndExportDocuments event.</summary>
    '     Function EndExportDocumentsEventHandler(ByVal sender As Object, ByRef e As ExportDocumentsEventArgs) As PlugInExecuteReturnArgs
    ' 
    '     '' Export Folder
    '     ''' <summary>Event handler for the BeginExportFolder event.</summary>
    '     Function BeginExportFolderEventHandler(ByVal sender As Object, ByRef e As ExportFolderEventArgs) As PlugInExecuteReturnArgs
    '     ''' <summary>Event handler for the EndExportFolder event.</summary>
    '     Function EndExportFolderEventHandler(ByVal sender As Object, ByRef e As ExportFolderEventArgs) As PlugInExecuteReturnArgs
    ' </Removed by: Ernie at: 9/29/2014-2:24:19 PM on machine: ERNIE-THINK>

    '' Transform Document
    ''' <summary>Event handler for the BeginTransform event.</summary>
    Function BeginTransformDocumentEventHandler(ByVal sender As Object, ByRef e As TransformDocumentEventArgs) As PlugInExecuteReturnArgs
    ''' <summary>Event handler for the EndTransform event.</summary>
    Function EndTransformDocumentEventHandler(ByVal sender As Object, ByRef e As TransformDocumentEventArgs) As PlugInExecuteReturnArgs

    '' Import Document
    ''' <summary>Event handler for the BeginImportDocument event.</summary>
    Function BeginImportDocumentEventHandler(ByVal sender As Object, ByRef e As ImportDocumentArgs) As PlugInExecuteReturnArgs
    ''' <summary>Event handler for the ImportDocumentError event.</summary>
    Function ImportDocumentErrorEventHandler(ByVal sender As Object, ByRef e As DocumentImportErrorEventArgs) As PlugInExecuteReturnArgs
    ''' <summary>Event handler for the EndImportDocument event.</summary>
    Function EndImportDocumentEventHandler(ByVal sender As Object, ByRef e As ImportDocumentArgs) As PlugInExecuteReturnArgs

    '' Import Record
    ''' <summary>Event handler for the BeginImportDocument event.</summary>
    Function BeginImportRecordEventHandler(ByVal sender As Object, ByRef e As ImportDocumentArgs) As PlugInExecuteReturnArgs
    ''' <summary>Event handler for the ImportRecordError event.</summary>
    Function ImportRecordErrorEventHandler(ByVal sender As Object, ByRef e As RecordImportErrorEventArgs) As PlugInExecuteReturnArgs
    ''' <summary>Event handler for the EndImportRecord event.</summary>
    Function EndImportRecordEventHandler(ByVal sender As Object, ByRef e As ImportDocumentArgs) As PlugInExecuteReturnArgs

    '' Migrate Document
    ''' <summary>Event handler for the BeginMigrateDocument event.</summary>
    Function BeginMigrateDocumentEventHandler(ByVal sender As Object, ByRef e As MigrateDocumentEventArgs) As PlugInExecuteReturnArgs
    ''' <summary>Event handler for the MigrateDocumentError event.</summary>
    Function MigrateDocumentErrorEventHandler(ByVal sender As Object, ByRef e As MigrateDocumentErrorEventArgs) As PlugInExecuteReturnArgs
    ''' <summary>Event handler for the EndMigrateDocument event.</summary>
    Function EndMigrateDocumentEventHandler(ByVal sender As Object, ByRef e As MigrateDocumentEventArgs) As PlugInExecuteReturnArgs

#End Region

  End Interface

End Namespace