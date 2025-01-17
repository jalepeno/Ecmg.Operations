'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports Documents
Imports Documents.Arguments
Imports Documents.Arguments.PlugInExecuteReturnArgs
Imports Documents.Utilities

#End Region

Namespace PlugIns

  ''' <summary>
  ''' Sub Class for Document Centric PlugIns
  ''' </summary>
  ''' <remarks></remarks>
  Public MustInherit Class CDocumentPlugIn
    Inherits CPlugIn
    Implements IDocumentPlugin

#Region "Private Properties"
    Private mstrDocumentPath As String = String.Empty
    Private mobjDocument As Core.Document = Nothing
#End Region

#Region "Constructors"
    Public Sub New()
      MyBase.New()
      Try
        Initialize()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub
#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property PlugInClassName() As String
      Get
        Return "CDocumentProcessingPlugIn"
      End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
      Get
        Return "Unknown"
      End Get
    End Property

    Public Overrides ReadOnly Property PlugInExecutionMode() As PlugInExecutionMode
      Get
        Return PlugIns.PlugInExecutionMode.Synchronous
      End Get
    End Property

    ''' <summary>
    ''' Gets the version of this plugin.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property Version() As String
      Get
        'Return "Unknown"
        Return Assembly.GetExecutingAssembly.ImageRuntimeVersion
      End Get
    End Property

    ''' <summary>
    ''' Path to the document the plugin will use to work with
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DocumentPath() As String
      Get
        Return mstrDocumentPath
      End Get
      Set(ByVal value As String)
        mstrDocumentPath = value
      End Set
    End Property

    ''' <summary>
    ''' Document object the plugin will use to work with
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Document() As Core.Document
      Get
        Return mobjDocument
      End Get
      Set(ByVal value As Core.Document)
        mobjDocument = value
      End Set
    End Property
#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Called immediately after document is added to repository
    ''' </summary>
    ''' <param name="lpDocument"></param>
    ''' <remarks></remarks>
    Public Overridable Function AfterDocumentAddedToRepository(ByVal lpDocument As Core.Document) As Arguments.PlugInExecuteReturnArgs
      Dim lobjReturn As New PlugInExecuteReturnArgs

      Try
        lobjReturn.ReturnCode = PlugInExecuteReturnArgs.PlugInExecuteReturnCode.Success
        lobjReturn.Document = lpDocument
      Catch ex As Exception
        'Don't throw exception here, just log it and let content loader move on
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

      Return lobjReturn
    End Function

    ''' <summary>
    ''' Called immediately after record is added to repository
    ''' </summary>
    ''' <param name="lpRecord"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function AfterRecordAddedToRepository(ByVal lpRecord As Object) As Arguments.PlugInExecuteReturnArgs
      Dim lobjReturn As New PlugInExecuteReturnArgs

      Try
        lobjReturn.ReturnCode = PlugInExecuteReturnArgs.PlugInExecuteReturnCode.Success
        lobjReturn.Record = lpRecord
      Catch ex As Exception
        'Don't throw exception here, just log it and let content loader move on
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

      Return lobjReturn
    End Function

    ''' <summary>
    ''' Called immediately before document is added to repository
    ''' This gives the plugin the chance to transform or manipulate the CDF before
    ''' it gets processed.
    ''' </summary>
    ''' <param name="lpDocument"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function BeforeDocumentAddedToRepository(ByVal lpDocument As Core.Document) As Arguments.PlugInExecuteReturnArgs
      Dim lobjReturn As New PlugInExecuteReturnArgs

      Try
        lobjReturn.ReturnCode = PlugInExecuteReturnArgs.PlugInExecuteReturnCode.Success
        lobjReturn.Document = lpDocument
      Catch ex As Exception
        'Don't throw exception here, just log it and let content loader move on
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

      Return lobjReturn
    End Function

    ''' <summary>
    ''' Called immediately before record is added to repository
    ''' This gives the plugin the chance to transform or manipulate the CDF before
    ''' it gets processed.
    ''' </summary>
    ''' <param name="lpRecord"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function BeforeRecordAddedToRepository(ByVal lpRecord As Object) As Arguments.PlugInExecuteReturnArgs
      Dim lobjReturn As New PlugInExecuteReturnArgs

      Try
        lobjReturn.ReturnCode = PlugInExecuteReturnArgs.PlugInExecuteReturnCode.Success
        lobjReturn.Record = lpRecord
      Catch ex As Exception
        'Don't throw exception here, just log it and let content loader move on
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

      Return lobjReturn
    End Function

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Initializes the plugin
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Initialize()
      Try
        AddProperties()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Throw
      End Try
    End Sub

    Private Sub AddProperties()
      ' Add the base properties here.
      Try
        'Add in any base properties here
        'PlugInProperties.Add(New PlugInProperty("NotifyAfterDocumentAdded", GetType(System.Boolean), True, "False"))
        'PlugInProperties.Add(New PlugInProperty("NotifyBeforeDocumentAdded", GetType(System.Boolean), True, "False"))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

    End Sub

#End Region

#Region "IDocumentPlugin Implementation"

    Public Overridable Function BeginProcessEventHandler(ByVal sender As Object, ByRef e As Arguments.DocumentEventArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.BeginProcessEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function BeginExportDocumentEventHandler(ByVal sender As Object, ByRef e As Arguments.ExportDocumentEventArgs) _
      As Arguments.PlugInExecuteReturnArgs _
      Implements IDocumentPlugin.BeginExportDocumentEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' <Removed by: Ernie at: 9/29/2014-2:25:47 PM on machine: ERNIE-THINK>
    '     Public Overridable Function BeginExportDocumentsEventHandler(ByVal sender As Object, ByRef e As Arguments.ExportDocumentsEventArgs) _
    '       As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.BeginExportDocumentsEventHandler
    '       Try
    '         Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success)
    '       Catch ex As Exception
    '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '         '  Re-throw the exception to the caller
    '         Throw
    '       End Try
    '     End Function
    ' 
    '     Public Overridable Function BeginExportFolderEventHandler(ByVal sender As Object, ByRef e As Arguments.ExportFolderEventArgs) _
    '       As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.BeginExportFolderEventHandler
    '       Try
    '         Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success)
    '       Catch ex As Exception
    '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '         '  Re-throw the exception to the caller
    '         Throw
    '       End Try
    '     End Function
    ' </Removed by: Ernie at: 9/29/2014-2:25:47 PM on machine: ERNIE-THINK>

    Public Overridable Function BeginImportDocumentEventHandler(ByVal sender As Object, ByRef e As Arguments.ImportDocumentArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.BeginImportDocumentEventHandler
      Try
        Dim lobjReturnArgs As New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document) With {
          .FolderToFileIn = e.FolderToFileIn
        }
        Return lobjReturnArgs
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function BeginImportRecordEventHandler(ByVal sender As Object, ByRef e As Arguments.ImportDocumentArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.BeginImportRecordEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function BeginMigrateDocumentEventHandler(ByVal sender As Object, ByRef e As Arguments.MigrateDocumentEventArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.BeginMigrateDocumentEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function BeginTransformDocumentEventHandler(ByVal sender As Object, ByRef e As Arguments.TransformDocumentEventArgs) _
 As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.BeginTransformDocumentEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function EndExportDocumentEventHandler(ByVal sender As Object, ByRef e As Arguments.ExportDocumentEventArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.EndExportDocumentEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' <Removed by: Ernie at: 9/29/2014-2:24:44 PM on machine: ERNIE-THINK>
    '     Public Overridable Function EndExportDocumentsEventHandler(ByVal sender As Object, ByRef e As Arguments.ExportDocumentsEventArgs) _
    '       As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.EndExportDocumentsEventHandler
    '       Try
    '         Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success)
    '       Catch ex As Exception
    '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '         '  Re-throw the exception to the caller
    '         Throw
    '       End Try
    '     End Function
    ' 
    '     Public Overridable Function EndExportFolderEventHandler(ByVal sender As Object, ByRef e As Arguments.ExportFolderEventArgs) _
    '       As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.EndExportFolderEventHandler
    '       Try
    '         Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success)
    '       Catch ex As Exception
    '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '         '  Re-throw the exception to the caller
    '         Throw
    '       End Try
    '     End Function
    ' </Removed by: Ernie at: 9/29/2014-2:24:44 PM on machine: ERNIE-THINK>

    Public Overridable Function EndImportDocumentEventHandler(ByVal sender As Object, ByRef e As Arguments.ImportDocumentArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.EndImportDocumentEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function EndImportRecordEventHandler(ByVal sender As Object, ByRef e As Arguments.ImportDocumentArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.EndImportRecordEventHandler
      Try
        Dim lobjReturnArgs As New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document) With {
          .Record = e.Document
        }
        Return lobjReturnArgs
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function EndMigrateDocumentEventHandler(ByVal sender As Object, ByRef e As Arguments.MigrateDocumentEventArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.EndMigrateDocumentEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function EndProcessEventHandler(ByVal sender As Object, ByRef e As Arguments.DocumentEventArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.EndProcessEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function EndTransformDocumentEventHandler(ByVal sender As Object, ByRef e As Arguments.TransformDocumentEventArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.EndTransformDocumentEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success, e.Document)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function ExportDocumentErrorEventHandler(ByVal sender As Object, ByRef e As Arguments.DocumentExportErrorEventArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.ExportDocumentErrorEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function ImportDocumentErrorEventHandler(ByVal sender As Object, ByRef e As Arguments.DocumentImportErrorEventArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.ImportDocumentErrorEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function ImportRecordErrorEventHandler(ByVal sender As Object, ByRef e As Arguments.RecordImportErrorEventArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.ImportRecordErrorEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function MigrateDocumentErrorEventHandler(ByVal sender As Object, ByRef e As Arguments.MigrateDocumentErrorEventArgs) _
      As Arguments.PlugInExecuteReturnArgs Implements IDocumentPlugin.MigrateDocumentErrorEventHandler
      Try
        Return New PlugInExecuteReturnArgs(PlugInExecuteReturnCode.Success)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace