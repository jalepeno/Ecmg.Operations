' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  WorkItem.vb
'  Description :  [type_description_here]
'  Created     :  12/5/2011 2:37:37 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json


#End Region

Public Class WorkItem
  Implements IWorkItem
  Implements IDisposable
  Implements ISerialize
  Implements IXmlSerializable

#Region "Class Constants"

  Private Const WORKITEM_FILE_EXTENSION As String = "wif"

#End Region

#Region "Class Variables"

  Private mdatCreateDate As DateTime = DateTime.MinValue
  Private mstrSourceDocId As String = String.Empty
  Private mstrDestinationDocId As String = String.Empty
  Private mobjDocument As Document = Nothing
  Private mobjFolder As Folder = Nothing
  Private mobjObject As CustomObject = Nothing
  Private mdatStartTime As DateTime = DateTime.MinValue
  Private mdatFinishTime As DateTime = DateTime.MinValue
  Private mstrId As String = String.Empty
  Private mobjParent As IItemParent = Nothing
  Private mobjProcess As IOperable = Nothing
  Private mobjProcessResults As IProcessResult = Nothing
  Private mstrProcessedBy As String = String.Empty
  Private mstrProcessedMessage As String = String.Empty
  Private menuProcessedStatus As ProcessedStatus = OperationEnumerations.ProcessedStatus.NotProcessed
  Private mobjTag As Object = Nothing
  Private mstrTitle As String = String.Empty
  Private mstrOriginalFilePath As String = String.Empty

#End Region

#Region "Constructors"

  Public Sub New()
    mdatCreateDate = Now
  End Sub

  Public Sub New(lpSourceConnectionString As String, lpSourceDocId As String)
    Me.New(lpSourceConnectionString, lpSourceDocId, String.Empty)
  End Sub

  Public Sub New(lpSourceConnection As IRepositoryConnection, lpSourceDocId As String)
    Me.New(lpSourceConnection, lpSourceDocId, String.Empty)
  End Sub

  Public Sub New(lpSourceConnectionString As String, lpSourceDocId As String, lpTitle As String)
    Try
      Id = Guid.NewGuid.ToString
      Parent = New ItemParent(lpSourceConnectionString)
      SourceDocId = lpSourceDocId
      Title = lpTitle
      mdatCreateDate = Now
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(lpSourceConnection As IRepositoryConnection, lpSourceDocId As String, lpTitle As String)
    Try
      Id = Guid.NewGuid.ToString
      Parent = New ItemParent(lpSourceConnection)
      SourceDocId = lpSourceDocId
      Title = lpTitle
      mdatCreateDate = Now
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(lpParent As IItemParent, lpSourceDocId As String)
    Me.New(lpParent, lpSourceDocId, String.Empty)
  End Sub

  Public Sub New(lpParent As IItemParent, lpSourceDocId As String, lpTitle As String)
    Try
      Id = Guid.NewGuid.ToString
      Parent = lpParent
      SourceDocId = lpSourceDocId
      Title = lpTitle
      mdatCreateDate = Now
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Properties"

  Public ReadOnly Property OriginalFilePath() As String
    Get
      Return mstrOriginalFilePath
    End Get
  End Property

#End Region

#Region "Private Properties"

  Protected ReadOnly Property IsDisposed() As Boolean
    Get
      Return disposedValue
    End Get
  End Property

#End Region

#Region "IWorkItem Implementation"

  Public ReadOnly Property CreateDate As DateTime Implements IWorkItem.CreateDate
    Get
      Return mdatCreateDate
    End Get
  End Property

  Public Property DestinationDocId As String Implements IWorkItem.DestinationDocId
    Get
      Return mstrDestinationDocId
    End Get
    Set(value As String)
      mstrDestinationDocId = value
    End Set
  End Property

  Public Property Document As Document Implements IWorkItem.Document
    Get
      Return mobjDocument
    End Get
    Set(value As Document)
      mobjDocument = value
    End Set
  End Property

  Public Property Folder As Folder Implements IWorkItem.Folder
    Get
      Return mobjFolder
    End Get
    Set(value As Folder)
      mobjFolder = value
    End Set
  End Property

  Public Property [Object] As CustomObject Implements IWorkItem.Object
    Get
      Return mobjObject
    End Get
    Set(value As CustomObject)
      mobjObject = value
    End Set
  End Property

  Public Property FinishTime As Date Implements IWorkItem.FinishTime
    Get
      Return mdatFinishTime
    End Get
    Set(value As Date)
      mdatFinishTime = value
    End Set
  End Property

  Public Property Id As String Implements IWorkItem.Id
    Get
      Return mstrId
    End Get
    Set(value As String)
      mstrId = value
    End Set
  End Property

  Public Property Parent As IItemParent Implements IWorkItem.Parent
    Get
      Return mobjParent
    End Get
    Set(value As IItemParent)
      mobjParent = value
    End Set
  End Property

  ''' <summary>
  ''' Gets or sets the process associated with the work item.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Process As IOperable Implements IWorkItem.Process
    Get
      Return mobjProcess
    End Get
    Set(value As IOperable)
      mobjProcess = value
    End Set
  End Property

  Public Property ProcessResult As IProcessResult Implements IWorkItem.ProcessResult
    Get
      Return mobjProcessResults
    End Get
    Set(ByVal value As IProcessResult)
      mobjProcessResults = value
    End Set
  End Property

  Public Property ProcessedBy As String Implements IWorkItem.ProcessedBy
    Get
      Return mstrProcessedBy
    End Get
    Set(value As String)
      mstrProcessedBy = value
    End Set
  End Property

  Public Property ProcessedMessage As String Implements IWorkItem.ProcessedMessage
    Get
      Return mstrProcessedMessage
    End Get
    Set(value As String)
      mstrProcessedMessage = value
    End Set
  End Property

  Public Property ProcessedStatus As OperationEnumerations.ProcessedStatus Implements IWorkItem.ProcessedStatus
    Get
      Return menuProcessedStatus
    End Get
    Set(value As OperationEnumerations.ProcessedStatus)
      menuProcessedStatus = value
    End Set
  End Property

  Public Property SourceDocId As String Implements IWorkItem.SourceDocId
    Get
      Return mstrSourceDocId
    End Get
    Set(value As String)
      mstrSourceDocId = value
    End Set
  End Property

  Public Property StartTime As Date Implements IWorkItem.StartTime
    Get
      Return mdatStartTime
    End Get
    Set(value As Date)
      mdatStartTime = value
    End Set
  End Property

  Public Property Tag As Object Implements IWorkItem.Tag
    Get
      Return mobjTag
    End Get
    Set(value As Object)
      mobjTag = value
    End Set
  End Property

  Public Property Title As String Implements IWorkItem.Title
    Get
      Return mstrTitle
    End Get
    Set(value As String)
      mstrTitle = value
    End Set
  End Property

  Public Property TotalProcessingTime As System.TimeSpan Implements IWorkItem.TotalProcessingTime
    Get
      Return FinishTime - StartTime
    End Get
    Set(value As System.TimeSpan)
      ' Do nothing
    End Set
  End Property

  Public Overridable Function Execute(lpProcess As IOperable) As Boolean Implements IWorkItem.Execute
    Try

      Process = lpProcess

      Me.ProcessedBy = Environment.MachineName

      Select Case lpProcess.Execute(Me)
        Case Result.Success
          Me.ProcessedStatus = OperationEnumerations.ProcessedStatus.Success
          Return True
        Case Result.Failed
          Me.ProcessedStatus = OperationEnumerations.ProcessedStatus.Failed
          Return False
        Case Result.NotProcessed
          Me.ProcessedStatus = OperationEnumerations.ProcessedStatus.NotProcessed
      End Select

      Me.ProcessedStatus = OperationEnumerations.ProcessedStatus.Failed
      Return False

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Me.ProcessedStatus = OperationEnumerations.ProcessedStatus.Failed
      Return False
    End Try
  End Function

  Public Overridable Function ToJsonString(ByVal lpIncludeProcessResult As Boolean) As String Implements IWorkItem.ToJsonString
    Try
      ' We will do manual JSON serialization for complete speed and control

      Return ToJsonString(Me, lpIncludeProcessResult)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function ToJsonString(ByVal lpWorkItem As IWorkItem, ByVal lpIncludeProcessResult As Boolean) As String
    Try
      ' We will do manual JSON serialization for complete speed and control

      Dim lobjStringBuilder As New StringBuilder
      Dim lobjStringWriter As New StringWriter(lobjStringBuilder)

      Using lobjJSONWriter As New JsonTextWriter(lobjStringWriter)
        With lobjJSONWriter
          .Formatting = Newtonsoft.Json.Formatting.Indented

          .WriteRaw("{""WorkItem"": ")
          .WriteStartObject()

          .WritePropertyName("Id")
          .WriteValue(lpWorkItem.Id)

          .WritePropertyName("ParentId")
          If lpWorkItem.Parent IsNot Nothing Then
            .WriteValue(lpWorkItem.Parent.Id)
          Else
            .WriteNull()
          End If

          .WritePropertyName("Title")
          .WriteValue(lpWorkItem.Title)

          .WritePropertyName("SourceDocId")
          .WriteValue(lpWorkItem.SourceDocId)

          .WritePropertyName("DestinationDocId")
          .WriteValue(lpWorkItem.DestinationDocId)

          .WritePropertyName("ProcessedStatus")
          .WriteValue(lpWorkItem.ProcessedStatus.ToString)

          If lpIncludeProcessResult Then
            .WritePropertyName("ProcessResult")
            If lpWorkItem.ProcessResult IsNot Nothing Then
              .WriteRawValue(lpWorkItem.ProcessResult.ToJsonString)
            Else
              .WriteNull()
            End If
          End If

          .WritePropertyName("ProcessedMessage")
          .WriteValue(lpWorkItem.ProcessedMessage)

          .WritePropertyName("StartTime")
          .WriteValue(lpWorkItem.StartTime.ToString)

          .WritePropertyName("FinishTime")
          .WriteValue(lpWorkItem.FinishTime.ToString)

          .WritePropertyName("TotalProcessingTime")
          .WriteValue(lpWorkItem.TotalProcessingTime.ToString)

          .WritePropertyName("ProcessedBy")
          .WriteValue(lpWorkItem.ProcessedBy)

          .WritePropertyName("CreateDate")
          .WriteValue(lpWorkItem.CreateDate.ToString)

          .WriteEndObject()
          .WriteRaw("}")
        End With
      End Using

      Return lobjStringBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "ISerialize Implementation"

  Public ReadOnly Property DefaultFileExtension As String _
    Implements ISerialize.DefaultFileExtension
    Get
      Return WORKITEM_FILE_EXTENSION
    End Get
  End Property

  Public Function Deserialize(lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object _
    Implements ISerialize.Deserialize
    Try

      If IsDisposed Then
        Throw New ObjectDisposedException(Me.GetType.ToString)
      End If

      mstrOriginalFilePath = lpFilePath
      Dim lobjProcess As IWorkItem = CType(Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType), IWorkItem)

      Return lobjProcess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function Deserialize(lpXML As System.Xml.XmlDocument) As Object _
    Implements ISerialize.Deserialize
    Try

      If IsDisposed Then
        Throw New ObjectDisposedException(Me.GetType.ToString)
      End If

      Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function Serialize() As System.Xml.XmlDocument _
    Implements ISerialize.Serialize
    Try
      Return Helper.FormatXmlDocument(Serializer.Serialize.Xml(Me))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Sub Serialize(ByRef lpFilePath As String, lpFileExtension As String) _
    Implements ISerialize.Serialize
    Try

      If lpFileExtension.Length = 0 Then
        ' No override was provided
        If lpFilePath.EndsWith(DefaultFileExtension) = False Then
          lpFilePath = lpFilePath.Remove(lpFilePath.Length - 3) & DefaultFileExtension
        End If

      End If

      Serializer.Serialize.XmlFile(Me, lpFilePath)

      Helper.FormatXmlFile(lpFilePath)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub Serialize(lpFilePath As String) _
    Implements ISerialize.Serialize
    Try
      Serialize(lpFilePath, String.Empty)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub Serialize(lpFilePath As String, lpWriteProcessingInstruction As Boolean, lpStyleSheetPath As String) _
    Implements ISerialize.Serialize
    Try
      If lpWriteProcessingInstruction = True Then
        Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
      Else
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      End If

      Helper.FormatXmlFile(lpFilePath)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Function ToXmlString() As String _
    Implements ISerialize.ToXmlString
    Try
      Return Helper.FormatXmlString(Serializer.Serialize.XmlString(Me))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "IXmlSerializable Implementation"

  Public Function GetSchema() As System.Xml.Schema.XmlSchema _
    Implements System.Xml.Serialization.IXmlSerializable.GetSchema
    ' As per the Microsoft guidelines this is not implemented
    Return Nothing
  End Function

  Public Sub ReadXml(reader As System.Xml.XmlReader) _
    Implements System.Xml.Serialization.IXmlSerializable.ReadXml
    Try

      Dim lobjXmlDocument As New XmlDocument
      Dim lobjAttribute As XmlAttribute = Nothing

      lobjXmlDocument.Load(reader)

      With lobjXmlDocument

      End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub WriteXml(writer As System.Xml.XmlWriter) _
    Implements System.Xml.Serialization.IXmlSerializable.WriteXml
    Try

      With writer

        ' Write the Id attribute
        .WriteAttributeString("Id", Me.Id)

        ' Write the ParentId attribute
        If (Parent IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(Me.Parent.Id)) Then
          .WriteAttributeString("ParentId", Me.Parent.Id)
        End If

        ' Write the Title attribute
        .WriteAttributeString("Title", Me.Title)

        ' Write the SourceDocId attribute
        .WriteAttributeString("SourceDocId", Me.SourceDocId)

        ' Write the DestinationDocId attribute
        .WriteAttributeString("DestinationDocId", Me.DestinationDocId)

        ' Write the ProcessedStatus attribute
        .WriteAttributeString("ProcessedStatus", Me.ProcessedStatus.ToString)

        ' Write the result, if applicable
        .WriteAttributeString("ProcessedMessage", Me.ProcessedMessage)

        ' Write the StartTime attribute
        .WriteAttributeString("StartTime", Me.StartTime.ToString)

        ' Write the FinishTime attribute
        .WriteAttributeString("FinishTime", Me.FinishTime.ToString)

        ' Write the TotalProcessingTime attribute
        .WriteAttributeString("TotalProcessingTime", Me.TotalProcessingTime.ToString)

        ' Write the ProcessedBy attribute
        .WriteAttributeString("ProcessedBy", Me.ProcessedBy)

        ' Write the CreateDate attribute
        .WriteAttributeString("CreateDate", Helper.ToDetailedDateString(Me.CreateDate))

        ' Write the locale attribute
        .WriteAttributeString("Locale", CultureInfo.CurrentCulture.Name)

        ' Write the Process as a node
        If Me.Process IsNot Nothing AndAlso TypeOf Me.Process Is Process Then
          .WriteRaw(Serializer.Serialize.XmlElementString(Me.Process))
        End If

      End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        ' DISPOSETODO: dispose managed state (managed objects).
      End If

      ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' DISPOSETODO: set large fields to null.
    End If
    Me.disposedValue = True
  End Sub

  ' DISPOSETODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    Dispose(True)
    GC.SuppressFinalize(Me)
  End Sub

#End Region

End Class
