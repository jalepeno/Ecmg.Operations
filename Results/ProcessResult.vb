
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ProcessResults.vb
'  Description :  [type_description_here]
'  Created     :  12/8/2011 7:34:52 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Public Class ProcessResult
  Inherits OperationResult
  Implements IProcessResult
  Implements IXmlSerializable

#Region "Class Constants"

  Protected Friend Const NODE_RUN_BEFORE_JOB_BEGIN As String = "RunBeforeJobBegin"
  Protected Friend Const NODE_RUN_AFTER_JOB_COMPLETE As String = "RunAfterJobComplete"
  Protected Friend Const NODE_RUN_ON_JOB_FAILURE As String = "RunOnJobFailure"

#End Region

#Region "Class Variables"

  Private mstrNode As String
  Private mintContentCount As Integer
  Private menuTotalContentSize As FileSize
  Private mintVersionCount As Integer

  Private mobjOperationResults As IOperableResults = New OperationResults
  Private mobjRunBeforeJobBeginResults As IOperableResults = New OperationResults
  Private mobjRunAfterJobCompleteResults As IOperableResults = New OperationResults
  Private mobjRunOnJobFailureResults As IOperableResults = New OperationResults

#End Region

#Region "Public Properties"

  Public Property Node As String Implements IProcessResult.Node
    Get
      Try
        Return mstrNode
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As String)
      Try
        mstrNode = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property RunBeforeJobBeginResults As IOperableResults Implements IProcessResult.RunBeforeJobBeginResults
    Get
      Try
        Return mobjRunBeforeJobBeginResults
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperableResults)
      Try
        mobjRunBeforeJobBeginResults = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property RunAfterJobCompleteResults As IOperableResults Implements IProcessResult.RunAfterJobCompleteResults
    Get
      Try
        Return mobjRunAfterJobCompleteResults
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperableResults)
      Try
        mobjRunAfterJobCompleteResults = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property RunOnJobFailureResults As IOperableResults Implements IProcessResult.RunOnJobFailureResults
    Get
      Try
        Return mobjRunOnJobFailureResults
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperableResults)
      Try
        mobjRunOnJobFailureResults = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property OperationResults As IOperableResults Implements IProcessResult.OperationResults
    Get
      Return mobjOperationResults
    End Get
    Set(ByVal value As IOperableResults)
      mobjOperationResults = value
    End Set
  End Property

  Public Property ContentCount As Integer Implements IProcessResult.ContentCount
    Get
      Return mintContentCount
    End Get
    Set(value As Integer)
      mintContentCount = value
    End Set
  End Property

  Public Property TotalContentSize As FileSize Implements IProcessResult.TotalContentSize
    Get
      Return menuTotalContentSize
    End Get
    Set(value As FileSize)
      menuTotalContentSize = value
    End Set
  End Property

  Public Property VersionCount As Integer Implements IProcessResult.VersionCount
    Get
      Return mintVersionCount
    End Get
    Set(value As Integer)
      mintVersionCount = value
    End Set
  End Property

#End Region

#Region "Constructors"

  Public Sub New()
  End Sub

  Public Sub New(ByVal lpProcess As IProcess)
    Try
      AssignFromProcess(lpProcess)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Methods"

  'Public Overrides Function ToXmlElementString() As String
  '  Try
  '    Return Serializer.Serialize.XmlElementString(Me)
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    '  Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  Public Overrides Function ToJsonString() As String
    Try
      ' We will do manual JSON serialization for complete speed and control

      Dim lobjStringBuilder As New StringBuilder
      Dim lobjStringWriter As New StringWriter(lobjStringBuilder)

      Using lobjJSONWriter As New JsonTextWriter(lobjStringWriter)
        With lobjJSONWriter
          .Formatting = Newtonsoft.Json.Formatting.Indented

          .WriteRaw("{""ProcessResult"": ")

          .WriteStartObject()

          .WritePropertyName("Name")
          .WriteValue(Name)

          .WritePropertyName("Scope")
          .WriteValue(Scope.ToString)

          .WritePropertyName("Result")
          .WriteValue(Result.ToString)

          .WritePropertyName("ProcessedMessage")
          .WriteValue(ProcessedMessage)

          .WritePropertyName("StartTime")
          .WriteValue(StartTime.ToString)

          .WritePropertyName("FinishTime")
          .WriteValue(FinishTime.ToString)

          .WritePropertyName("TotalProcessingTime")
          .WriteValue(TotalProcessingTime.ToString)

          .WritePropertyName("WorkItem")
          If ((Parent IsNot Nothing) AndAlso (Parent.WorkItem IsNot Nothing)) Then
            .WriteRawValue(Parent.WorkItem.ToJsonString(False))
          Else
            .WriteNull()
          End If

          .WritePropertyName("OperationResults")
          .WriteStartArray()
          For Each lobjResult As IOperableResult In OperationResults
            .WriteRawValue(lobjResult.ToJsonString)
          Next
          .WriteEndArray()

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

#Region "Private Methods"

  Private Sub AssignFromProcess(ByVal lpProcess As IProcess)
    Try

      SetParent(lpProcess)

      With lpProcess
        Me.Name = .Name
        Me.Node = System.Net.Dns.GetHostName()
        Me.Result = .Result
        Me.ProcessedMessage = .ProcessedMessage
        Me.StartTime = .StartTime
        Me.FinishTime = .FinishTime
        Me.TotalProcessingTime = .TotalProcessingTime
      End With

      If lpProcess.WorkItem IsNot Nothing AndAlso lpProcess.WorkItem.Document IsNot Nothing Then
        With lpProcess.WorkItem.Document
          Me.ContentCount = .ContentCount
          Me.VersionCount = .Versions.Count
          Me.TotalContentSize = .TotalContentSize
        End With
      End If

      For Each lobjOperation As IOperable In lpProcess.Operations
        Me.OperationResults.Add(New OperationResult(lobjOperation))
      Next

      ' Assign the results from all of the event operations
      SetEventResults(lpProcess.RunBeforeJobBegin, Me.RunBeforeJobBeginResults)
      SetEventResults(lpProcess.RunAfterJobComplete, Me.RunAfterJobCompleteResults)
      SetEventResults(lpProcess.RunOnJobFailure, Me.RunOnJobFailureResults)
      SetEventResults(lpProcess.RunBeforeParentBegin, Me.RunBeforeParentBeginResults)
      SetEventResults(lpProcess.RunAfterParentComplete, Me.RunAfterParentCompleteResults)
      SetEventResults(lpProcess.RunOnParentFailure, Me.RunOnParentFailureResults)
      SetEventResults(lpProcess.RunBeforeBegin, Me.RunBeforeBeginResults)
      SetEventResults(lpProcess.RunAfterComplete, Me.RunAfterCompleteResults)
      SetEventResults(lpProcess.RunOnFailure, Me.RunOnFailureResults)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Friend Overrides Function DebuggerIdentifier() As String
    Return MyBase.DebuggerIdentifier()
  End Function

  Private Shared Sub SetEventResults(ByRef lpEventOperable As IOperable, ByRef lpEventOperableResults As IOperableResults)
    Try

      If lpEventOperable IsNot Nothing Then
        lpEventOperableResults = New OperationResults
        If TypeOf (lpEventOperable) Is IOperation Then
          lpEventOperableResults.Add(New OperationResult(lpEventOperable))
        ElseIf TypeOf (lpEventOperable) Is IProcess Then
          For Each lobjOperation As IOperable In CType(lpEventOperable, IProcess).Operations
            lpEventOperableResults.Add(New OperationResult(lpEventOperable))
          Next
        End If
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "IXmlSerializable Implementation"

  Public Overrides Sub ReadXml(ByVal reader As System.Xml.XmlReader)

    Dim lobjXmlDocument As New XmlDocument

    Try

      'Dim lstrCurrentElementName As String
      'Dim lobjOperationResult As OperationResult = Nothing

      'MyBase.ReadXml(reader)

      'Do Until reader.NodeType = XmlNodeType.EndElement AndAlso reader.Name = "ProcessResults"

      '  If reader.NodeType = XmlNodeType.Element Then
      '    lstrCurrentElementName = reader.Name
      '    If lstrCurrentElementName = "OperationResult" Then
      '      lobjOperationResult = New OperationResult
      '      lobjOperationResult.ReadXml(reader)
      '      OperationResults.Add(lobjOperationResult)
      '    End If

      '  End If

      '  reader.Read()
      '  Do Until reader.NodeType <> XmlNodeType.Whitespace
      '    reader.Read()
      '  Loop

      'Loop

      lobjXmlDocument.Load(reader)

      With lobjXmlDocument
        OperationResult.ReadResultsXml(Me, lobjXmlDocument.DocumentElement)
      End With

      Dim lobjChildOperationNodes As XmlNodeList = lobjXmlDocument.SelectNodes("//ChildOperations/OperationResult")

      If lobjChildOperationNodes IsNot Nothing Then
        Dim lobjChildOperationResult As IOperableResult = Nothing
        For Each lobjChildResultNode As XmlElement In lobjChildOperationNodes
          If String.IsNullOrEmpty(lobjChildResultNode.InnerText) Then
            Continue For
          End If
          lobjChildOperationResult = New OperationResult
          OperationResult.ReadResultsXml(lobjChildOperationResult, lobjChildResultNode)
          Me.ChildOperations.Add(lobjChildOperationResult)
        Next
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overrides Sub WriteXml(ByVal writer As System.Xml.XmlWriter)
    Try

      With writer

        .WriteAttributeString("Name", Name)
        .WriteAttributeString("Node", Node)
        .WriteAttributeString("Result", Result.ToString)
        .WriteAttributeString("ProcessedMessage", ProcessedMessage)
        .WriteAttributeString("StartTime", StartTime.ToString)
        .WriteAttributeString("FinishTime", FinishTime.ToString)
        .WriteAttributeString("TotalProcessingTime", TotalProcessingTime.ToString)
        .WriteAttributeString("VersionCount", VersionCount.ToString)
        .WriteAttributeString("ContentCount", ContentCount.ToString)
        If TotalContentSize IsNot Nothing Then
          .WriteAttributeString("TotalContentSize", TotalContentSize.ToString())
        Else
          .WriteAttributeString("TotalContentSize", New FileSize().ToString())
        End If


        WritePreOperationEventResults(writer)

        .WriteStartElement("OperationResults")

        For Each lobjOperationResult As IOperableResult In OperationResults
          If lobjOperationResult IsNot Nothing Then
            .WriteRaw(lobjOperationResult.ToXmlElementString)
          End If
        Next

        .WriteEndElement()

        .WriteStartElement("ChildOperations")
        For Each lobjChildOperationResult As IOperableResult In ChildOperations
          .WriteRaw(lobjChildOperationResult.ToXmlElementString)
        Next
        .WriteEndElement()

      End With

      WritePostOperationEventResults(writer)
      WriteOnFailureEventResults(writer)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overrides Sub WritePreOperationEventResults(writer As System.Xml.XmlWriter)
    Try

      'With writer

      WriteEventResults(writer, RunBeforeJobBeginResults, NODE_RUN_BEFORE_JOB_BEGIN)

      '  If RunBeforeJobBeginResults IsNot Nothing Then
      '    ' Write the RunBeforeJobBegin result
      '    ' Open the RunBeforeJobBegin Element
      '    .WriteStartElement("RunBeforeJobBegin")

      '    .WriteRaw(RunBeforeJobBeginResults.ToXmlElementString)

      '    ' End the element
      '    .WriteEndElement()
      '  End If

      'End With

      MyBase.WritePreOperationEventResults(writer)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overrides Sub WritePostOperationEventResults(writer As System.Xml.XmlWriter)
    Try

      MyBase.WritePostOperationEventResults(writer)
      WriteEventResults(writer, RunAfterJobCompleteResults, NODE_RUN_AFTER_JOB_COMPLETE)

      'With writer

      '  If RunAfterJobCompleteResults IsNot Nothing Then
      '    ' Write the RunAfterJobComplete result
      '    ' Open the RunAfterJobComplete Element
      '    .WriteStartElement("RunAfterJobComplete")

      '    .WriteRaw(RunAfterJobCompleteResults.ToXmlElementString)

      '    ' End the element
      '    .WriteEndElement()
      '  End If

      'End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overrides Sub WriteOnFailureEventResults(writer As System.Xml.XmlWriter)
    Try
      MyBase.WriteOnFailureEventResults(writer)
      WriteEventResults(writer, RunOnJobFailureResults, NODE_RUN_ON_JOB_FAILURE)

      'With writer

      '  If RunOnJobFailureResults IsNot Nothing Then
      '    ' Write the RunOnJobFailure result
      '    ' Open the RunOnJobFailure Element
      '    .WriteStartElement("RunOnJobFailure")

      '    .WriteRaw(RunOnJobFailureResults.ToXmlElementString)

      '    ' End the element
      '    .WriteEndElement()
      '  End If

      'End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region
End Class
