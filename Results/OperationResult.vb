' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperationResult.vb
'  Description :  [type_description_here]
'  Created     :  12/8/2011 7:38:01 AM
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
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

<DebuggerDisplay("{DebuggerIdentifier(),nq}")>
Public Class OperationResult
  Implements IOperableResult
  Implements IXmlSerializable

#Region "Class Constants"

  Protected Const NODE_RUN_BEFORE_BEGIN As String = "RunBeforeBegin"
  Protected Const NODE_RUN_AFTER_COMPLETE As String = "RunAfterComplete"
  Protected Const NODE_RUN_ON_FAILURE As String = "RunOnFailure"
  Protected Const NODE_RUN_BEFORE_PARENT_BEGIN As String = "RunBeforeParentBegin"
  Protected Const NODE_RUN_AFTER_PARENT_COMPLETE As String = "RunAfterParentComplete"
  Protected Const NODE_RUN_ON_PARENT_FAILURE As String = "RunOnParentFailure"

#End Region

#Region "Class Variables"

  Private mstrName As String = String.Empty
  Private mobjParent As IOperable = Nothing
  Private menuResult As OperationEnumerations.Result = Result.NotProcessed
  Private mstrProcessedMessage As String = String.Empty
  Private menuScope As OperationScope = OperationScope.Source
  Private mdatStartTime As DateTime = DateTime.MinValue
  Private mdatFinishTime As DateTime = DateTime.MinValue
  Private mobjTotalProcessingTime As TimeSpan = TimeSpan.Zero
  Private mobjChildOperations As IOperableResults = New OperationResults

  Private mobjRunBeforeBeginResults As IOperableResults = New OperationResults
  Private mobjRunAfterCompleteResults As IOperableResults = New OperationResults
  Private mobjRunOnFailureResults As IOperableResults = New OperationResults
  Private mobjRunBeforeParentBeginResults As IOperableResults = New OperationResults
  Private mobjRunAfterParentCompleteResults As IOperableResults = New OperationResults
  Private mobjRunOnParentFailureResults As IOperableResults = New OperationResults

#End Region

#Region "Public Properties"

  Public Property Name As String Implements IOperableResult.Name
    Get
      Return mstrName
    End Get
    Set(ByVal value As String)
      mstrName = value
    End Set
  End Property

  ReadOnly Property Parent As IOperable Implements IOperableResult.Parent
    Get
      Return mobjParent
    End Get
  End Property

  Public Property Result As OperationEnumerations.Result Implements IOperableResult.Result
    Get
      Return menuResult
    End Get
    Set(ByVal value As OperationEnumerations.Result)
      menuResult = value
    End Set
  End Property

  Public Property ProcessedMessage As String Implements IOperableResult.ProcessedMessage
    Get
      Return mstrProcessedMessage
    End Get
    Set(ByVal value As String)
      mstrProcessedMessage = value
    End Set
  End Property

  Public Property Scope As OperationScope Implements IOperableResult.Scope
    Get
      Return menuScope
    End Get
    Set(ByVal value As OperationScope)
      menuScope = value
    End Set
  End Property

  Public Property StartTime As DateTime Implements IOperableResult.StartTime
    Get
      Return mdatStartTime
    End Get
    Set(ByVal value As DateTime)
      mdatStartTime = value
    End Set
  End Property

  Public Property FinishTime As DateTime Implements IOperableResult.FinishTime
    Get
      Return mdatFinishTime
    End Get
    Set(ByVal value As DateTime)
      mdatFinishTime = value
    End Set
  End Property

  Public Property TotalProcessingTime As TimeSpan Implements IOperableResult.TotalProcessingTime
    Get
      Return mobjTotalProcessingTime
    End Get
    Set(ByVal value As TimeSpan)
      mobjTotalProcessingTime = value
    End Set
  End Property

  Public Property ChildOperations As IOperableResults Implements IOperableResult.ChildOperations
    Get
      Return mobjChildOperations
    End Get
    Set(ByVal value As IOperableResults)
      mobjChildOperations = value
    End Set
  End Property

  Public Property RunBeforeBeginResults As IOperableResults Implements IOperableResult.RunBeforeBeginResults
    Get
      Try
        Return mobjRunBeforeBeginResults
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperableResults)
      Try
        mobjRunBeforeBeginResults = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property RunAfterCompleteResults As IOperableResults Implements IOperableResult.RunAfterCompleteResults
    Get
      Try
        Return mobjRunAfterCompleteResults
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperableResults)
      Try
        mobjRunAfterCompleteResults = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property RunOnFailureResults As IOperableResults Implements IOperableResult.RunOnFailureResults
    Get
      Try
        Return mobjRunOnFailureResults
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperableResults)
      Try
        mobjRunOnFailureResults = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property RunBeforeParentBeginResults As IOperableResults Implements IOperableResult.RunBeforeParentBeginResults
    Get
      Try
        Return mobjRunBeforeParentBeginResults
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperableResults)
      Try
        mobjRunBeforeParentBeginResults = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property RunAfterParentCompleteResults As IOperableResults Implements IOperableResult.RunAfterParentCompleteResults
    Get
      Try
        Return mobjRunAfterParentCompleteResults
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperableResults)
      Try
        mobjRunAfterParentCompleteResults = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property RunOnParentFailureResults As IOperableResults Implements IOperableResult.RunOnParentFailureResults
    Get
      Try
        Return mobjRunOnParentFailureResults
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperableResults)
      Try
        mobjRunOnParentFailureResults = value
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

  End Sub

  Public Sub New(ByVal lpOperation As IOperable)
    Try
      AssignFromOperation(lpOperation)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(ByVal lpXMLElement As XmlElement)
    Try
      ReadResultsXml(Me, lpXMLElement)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Protected Methods"

  Protected Friend Overridable Function DebuggerIdentifier() As String
    Dim lobjIdentifierBuilder As New Text.StringBuilder
    Try

      If Not String.IsNullOrEmpty(Name) Then
        lobjIdentifierBuilder.AppendFormat("{0}", Name)
      End If

      lobjIdentifierBuilder.AppendFormat(": {0}", Result.ToString)

      If Result <> OperationEnumerations.Result.NotProcessed Then
        lobjIdentifierBuilder.AppendFormat(" - ProcessingTime {0}", TotalProcessingTime.ToString)
      End If

      Return lobjIdentifierBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Return lobjIdentifierBuilder.ToString
    End Try
  End Function

  Protected Sub SetParent(lpParent As IOperable)
    Try
      mobjParent = lpParent
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Methods"

  Public Overridable Function ToJsonString() As String Implements IOperableResult.ToJsonString
    Try
      'Return Serializer.Serialize.JsonString(Me)
      ' We will do manual JSON serialization for complete speed and control

      Dim lobjStringBuilder As New StringBuilder
      Dim lobjStringWriter As New StringWriter(lobjStringBuilder)

      Using lobjJSONWriter As New JsonTextWriter(lobjStringWriter)
        With lobjJSONWriter
          .Formatting = Newtonsoft.Json.Formatting.Indented

          .WriteRaw("{""OperationResult"": ")

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

          .WriteEndObject()

          .WriteRaw("}")

        End With
      End Using

      Return lobjStringBuilder.ToString

    Catch Ex As Exception
      ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overridable Function ToXmlString() As String Implements IOperableResult.ToXmlString
    Try
      Return Serializer.Serialize.XmlString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overridable Function ToXmlElementString() As String Implements IOperableResult.ToXmlElementString
    Try
      Return Serializer.Serialize.XmlElementString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Sub AssignFromOperation(ByVal lpOperation As IOperable)
    Try

      SetParent(lpOperation)

      With lpOperation
        Me.Name = .Name
        Me.Result = .Result
        Me.ProcessedMessage = .ProcessedMessage

        If TypeOf lpOperation Is IOperation Then
          Me.Scope = CType(lpOperation, IOperation).Scope
        Else
          Me.Scope = OperationScope.Source
        End If

        Me.StartTime = .StartTime
        Me.FinishTime = .FinishTime
        Me.TotalProcessingTime = .TotalProcessingTime
      End With

      If lpOperation.RunBeforeBegin IsNot Nothing Then
        Me.ChildOperations.Add(New OperationResult(lpOperation.RunBeforeBegin))
      End If

      If lpOperation.RunAfterComplete IsNot Nothing Then
        Me.ChildOperations.Add(New OperationResult(lpOperation.RunAfterComplete))
      End If

      If lpOperation.RunOnFailure IsNot Nothing Then
        Me.ChildOperations.Add(New OperationResult(lpOperation.RunOnFailure))
      End If

      If TypeOf lpOperation Is IDecisionOperation Then
        Me.ChildOperations = New OperationResults
        For Each lobjChildOperation As IOperable In CType(lpOperation, IDecisionOperation).RunOperations
          Me.ChildOperations.Add(New OperationResult(lobjChildOperation))
        Next
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "IXmlSerializable Implementation"

  Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
    Return Nothing
  End Function

  Public Shared Sub ReadResultsXml(ByRef sender As IOperableResult, ByVal element As XmlElement)

    Try
      Dim lstrScope As String = Nothing

      With element
        sender.Name = .GetAttribute("Name")
        lstrScope = .GetAttribute("Scope")
        If Not String.IsNullOrEmpty(lstrScope) Then
          sender.Scope = CType([Enum].Parse(GetType(OperationScope), lstrScope), OperationScope)
        End If
        Dim lstrResult As String = .GetAttribute("Result")
        If Not String.IsNullOrEmpty(lstrResult) Then
          sender.Result = CType([Enum].Parse(GetType(Result), lstrResult), OperationEnumerations.Result)
        Else
          sender.Result = OperationEnumerations.Result.NotProcessed
        End If

        sender.ProcessedMessage = .GetAttribute("ProcessedMessage")
        Dim v = DateTime.TryParse(.GetAttribute("StartTime"), sender.StartTime)
        v = DateTime.TryParse(.GetAttribute("FinishTime"), sender.FinishTime)
        v = TimeSpan.TryParse(.GetAttribute("TotalProcessingTime"), sender.TotalProcessingTime)

        If TypeOf sender Is IProcessResult Then
          With DirectCast(sender, IProcessResult)
            .Node = element.GetAttribute("Node")
            v = Integer.TryParse(element.GetAttribute("VersionCount"), .VersionCount)
            v = Integer.TryParse(element.GetAttribute("ContentCount"), .ContentCount)
            Dim lstrTotalContentSize As String = element.GetAttribute("TotalContentSize")
            If Not String.IsNullOrEmpty(lstrTotalContentSize) Then
              .TotalContentSize = FileSize.FromString(lstrTotalContentSize)
            End If
          End With
        End If

        If TypeOf sender IsNot IProcessResult Then



          Dim lobjChildOperationsNode As XmlNode = .SelectSingleNode("ChildOperations")
          If lobjChildOperationsNode IsNot Nothing AndAlso lobjChildOperationsNode.HasChildNodes Then
            sender.ChildOperations = New OperationResults(lobjChildOperationsNode)
          Else
            sender.ChildOperations = New OperationResults
          End If
        Else
          Dim lobjOperationResultssNode As XmlNode = .SelectSingleNode("OperationResults")
          CType(sender, IProcessResult).OperationResults = New OperationResults(lobjOperationResultssNode)
        End If



      End With

      If Not Helper.CallStackContainsMethodName("ReadAllEventResults") Then
        ReadAllEventResults(sender, element)
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Sub

  Protected Shared Sub ReadAllEventResults(ByRef sender As IOperableResult, ByVal element As XmlElement)
    Try

      With sender
        .RunBeforeBeginResults = ReadEventResults(NODE_RUN_BEFORE_BEGIN, element)
        .RunAfterCompleteResults = ReadEventResults(NODE_RUN_AFTER_COMPLETE, element)
        .RunOnFailureResults = ReadEventResults(NODE_RUN_ON_FAILURE, element)
        .RunBeforeParentBeginResults = ReadEventResults(NODE_RUN_BEFORE_PARENT_BEGIN, element)
        .RunAfterParentCompleteResults = ReadEventResults(NODE_RUN_AFTER_PARENT_COMPLETE, element)
        .RunOnParentFailureResults = ReadEventResults(NODE_RUN_ON_PARENT_FAILURE, element)
      End With

      If TypeOf sender Is IProcessResult Then
        With CType(sender, IProcessResult)
          .RunBeforeJobBeginResults = ReadEventResults(ProcessResult.NODE_RUN_BEFORE_JOB_BEGIN, element)
          .RunAfterJobCompleteResults = ReadEventResults(ProcessResult.NODE_RUN_AFTER_JOB_COMPLETE, element)
          .RunOnJobFailureResults = ReadEventResults(ProcessResult.NODE_RUN_ON_JOB_FAILURE, element)
        End With
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Shared Function ReadEventResults(ByVal lpNodeName As String, ByVal lpElement As XmlElement) As IOperableResults
    Try

      Dim lobjEventNode As XmlNode = lpElement.SelectSingleNode(String.Format("//{0}", lpNodeName))
      If ((lobjEventNode IsNot Nothing) AndAlso (lobjEventNode.HasChildNodes = True)) Then
        Return New OperationResults(lobjEventNode.FirstChild)
      Else
        Return Nothing
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overridable Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
    Try

      Dim lstrScope As String = Nothing
      Dim lstrCurrentElementName As String = String.Empty

      With reader
        Name = .GetAttribute("Name")
        lstrScope = .GetAttribute("Scope")
        If Not String.IsNullOrEmpty(lstrScope) Then
          Scope = CType([Enum].Parse(GetType(OperationScope), lstrScope), OperationScope)
        End If
        Result = CType([Enum].Parse(GetType(Result), .GetAttribute("Result")), OperationEnumerations.Result)
        ProcessedMessage = .GetAttribute("ProcessedMessage")
        Dim v = DateTime.TryParse(.GetAttribute("StartTime"), StartTime)
        v = DateTime.TryParse(.GetAttribute("FinishTime"), FinishTime)
        v = TimeSpan.TryParse(.GetAttribute("TotalProcessingTime"), TotalProcessingTime)

        If TypeOf Me IsNot ProcessResult Then
          Do Until reader.NodeType = XmlNodeType.EndElement AndAlso (reader.Name.EndsWith("OperationResult") OrElse reader.Name = "ProcessResults")
            If reader.NodeType = XmlNodeType.Element Then
              lstrCurrentElementName = reader.Name
            Else
              Select Case lstrCurrentElementName
                Case "ChildOperations"

              End Select
            End If
            reader.Read()
          Loop
        End If

      End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overridable Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
    Try

      With writer

        .WriteAttributeString("Name", Name)
        .WriteAttributeString("Scope", Scope.ToString)
        .WriteAttributeString("Result", Result.ToString)
        .WriteAttributeString("ProcessedMessage", ProcessedMessage)
        .WriteAttributeString("StartTime", StartTime.ToString)
        .WriteAttributeString("FinishTime", FinishTime.ToString)
        .WriteAttributeString("TotalProcessingTime", TotalProcessingTime.ToString)

        WritePreOperationEventResults(writer)

        If ChildOperations IsNot Nothing Then
          ' Write the ChildOperations
          ' Open the ChildOperations Element
          .WriteStartElement("ChildOperations")

          For Each lobjChildOperation As IOperableResult In ChildOperations
            ' Write the Parameter element
            .WriteRaw(lobjChildOperation.ToXmlElementString)
          Next

          ' End the ChildOperations element
          .WriteEndElement()
        End If

      End With

      WritePostOperationEventResults(writer)
      WriteOnFailureEventResults(writer)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overridable Sub WritePreOperationEventResults(ByVal writer As System.Xml.XmlWriter)
    Try

      'With writer

      WriteEventResults(writer, RunBeforeParentBeginResults, NODE_RUN_BEFORE_PARENT_BEGIN)
      WriteEventResults(writer, RunBeforeBeginResults, NODE_RUN_BEFORE_BEGIN)

      'If RunBeforeParentBeginResults IsNot Nothing Then
      '  ' Write the RunBeforeParent result
      '  ' Open the RunBeforeParent Element
      '  .WriteStartElement(NODE_RUN_BEFORE_PARENT_BEGIN)

      '  .WriteRaw(RunBeforeParentBeginResults.ToXmlElementString)

      '  ' End the element
      '  .WriteEndElement()
      'End If

      'If RunBeforeBeginResults IsNot Nothing Then
      '  ' Write the RunBeforeBegin result
      '  ' Open the RunBeforeBegin Element
      '  .WriteStartElement(NODE_RUN_BEFORE_BEGIN)

      '  .WriteRaw(RunBeforeBeginResults.ToXmlElementString)

      '  ' End the element
      '  .WriteEndElement()
      'End If

      'End With
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overridable Sub WritePostOperationEventResults(ByVal writer As System.Xml.XmlWriter)
    Try

      ' With writer

      WriteEventResults(writer, RunAfterCompleteResults, NODE_RUN_AFTER_COMPLETE)
      WriteEventResults(writer, RunAfterParentCompleteResults, NODE_RUN_AFTER_PARENT_COMPLETE)

      'If RunAfterCompleteResults IsNot Nothing Then
      '  ' Write the RunAfterComplete result
      '  ' Open the RunAfterComplete Element
      '  .WriteStartElement(NODE_RUN_AFTER_COMPLETE)

      '  .WriteRaw(RunAfterCompleteResults.ToXmlElementString)

      '  ' End the element
      '  .WriteEndElement()
      'End If

      'If RunAfterParentCompleteResults IsNot Nothing Then
      '  ' Write the RunAfterParentComplete result
      '  ' Open the RunAfterParentComplete Element
      '  .WriteStartElement(NODE_RUN_AFTER_PARENT_COMPLETE)

      '  .WriteRaw(RunAfterParentCompleteResults.ToXmlElementString)

      '  ' End the element
      '  .WriteEndElement()
      'End If

      'End With
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Sub WriteEventResults(ByVal writer As System.Xml.XmlWriter, lpOperableResults As IOperableResults, lpElementName As String)
    Try

      With writer

        If ((lpOperableResults IsNot Nothing) AndAlso (lpOperableResults.Count > 0)) Then
          ' Write the result

          Dim lstrRawResultsXml As String = lpOperableResults.ToXmlElementString

          If Not String.IsNullOrEmpty(lstrRawResultsXml) Then
            lstrRawResultsXml = lstrRawResultsXml.Replace("ArrayOfOperationResult", "EventOperationResults")
          End If

          ' Open the Element
          .WriteStartElement(lpElementName)

          ' Write the value
          .WriteRaw(lstrRawResultsXml)
          '.WriteRaw(lpOperableResults.ToXmlElementString.Replace("ArrayOfOperationResult", ""))

          ' End the element
          .WriteEndElement()
        End If

      End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overridable Sub WriteOnFailureEventResults(ByVal writer As System.Xml.XmlWriter)
    Try

      'With writer

      WriteEventResults(writer, RunOnFailureResults, NODE_RUN_ON_FAILURE)
      WriteEventResults(writer, RunOnParentFailureResults, NODE_RUN_ON_PARENT_FAILURE)

      'If RunOnFailureResults IsNot Nothing Then
      '  ' Write the RunOnFailure result
      '  ' Open the RunOnFailure Element
      '  .WriteStartElement(NODE_RUN_ON_FAILURE)

      '  .WriteRaw(RunOnFailureResults.ToXmlElementString)

      '  ' End the element
      '  .WriteEndElement()
      'End If

      'If RunOnParentFailureResults IsNot Nothing Then
      '  ' Write the RunOnParentFailure result
      '  ' Open the RunOnParentFailure Element
      '  .WriteStartElement(NODE_RUN_ON_PARENT_FAILURE)

      '  .WriteRaw(RunOnParentFailureResults.ToXmlElementString)

      '  ' End the element
      '  .WriteEndElement()
      'End If

      'End With
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
    Try
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: dispose managed state (managed objects).
          mstrName = String.Empty
          mobjParent = Nothing
          menuResult = Nothing
          mstrProcessedMessage = Nothing
          menuScope = Nothing
          mdatStartTime = Nothing
          mdatFinishTime = Nothing
          mobjTotalProcessingTime = Nothing
          mobjChildOperations.Dispose()
        End If

        ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' DISPOSETODO: set large fields to null.
      End If
      Me.disposedValue = True
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  ' DISPOSETODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    GC.SuppressFinalize(Me)
  End Sub
#End Region

End Class
