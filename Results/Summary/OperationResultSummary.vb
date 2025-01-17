' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperationResultSummary.vb
'  Description :  [type_description_here]
'  Created     :  4/26/2016 10:26:40 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.Serialization
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

<DebuggerDisplay("{DebuggerIdentifier(),nq}")>
Public Class OperationResultSummary
  Implements IOperableResultSummary
  Implements IXmlSerializable
  Implements IDisposable

#Region "Class Variables"

  Private mstrName As String = String.Empty
  Private mdatCreateDate As DateTime = Nothing
  Private mstrNode As String = String.Empty
  Private mobjParent As IOperable = Nothing
  Private menuResult As OperationEnumerations.Result = Result.NotProcessed
  Private menuScope As OperationScope = OperationScope.Source
  Private mobjProcessingTime As New ProcessingTimeInfo

#End Region

#Region "Constructors"

  Protected Sub New()
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

  Public Sub New(ByVal lpResult As IOperableResult)
    Try

      With lpResult
        Me.Name = .Name
        Me.Result = .Result
        Me.Scope = .Scope
        SetParent(.Parent)
        Me.ProcessingTime = New ProcessingTimeInfo(.TotalProcessingTime)
      End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Properties"

  Public Property Name As String Implements IOperableResultSummary.Name
    Get
      Try
        Return mstrName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As String)
      Try
        mstrName = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property CreateDate As DateTime Implements IOperableResultSummary.CreateDate
    Get
      Try
        Return mdatCreateDate
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As DateTime)
      Try
        mdatCreateDate = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Node As String Implements IOperableResultSummary.Node
    Get
      Try
        Return mstrNode
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As String)
      Try
        mstrNode = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public ReadOnly Property Parent As IOperable Implements IOperableResultSummary.Parent
    Get
      'Throw New NotImplementedException()
      Return Nothing
    End Get
  End Property

  Public Property Result As Result Implements IOperableResultSummary.Result
    Get
      Try
        Return menuResult
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Result)
      Try
        menuResult = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Scope As OperationScope Implements IOperableResultSummary.Scope
    Get
      Try
        Return menuScope
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As OperationScope)
      Try
        menuScope = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property ProcessingTime As ProcessingTimeInfo Implements IOperableResultSummary.ProcessingTime
    Get
      Try
        Return mobjProcessingTime
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As ProcessingTimeInfo)
      Try
        mobjProcessingTime = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

#End Region

#Region "Protected Methods"

  Protected Friend Overridable Function DebuggerIdentifier() As String
    Dim lobjIdentifierBuilder As New Text.StringBuilder
    Try

      If Not String.IsNullOrEmpty(Name) Then
        lobjIdentifierBuilder.AppendFormat("{0}", Name)
      End If

      If TypeOf Me Is IProcessResultSummary Then
        lobjIdentifierBuilder.AppendFormat(" ({0} Operations)", DirectCast(Me, IProcessResultSummary).OperationResults.Count())
      End If
      'lobjIdentifierBuilder.AppendFormat(": {0}", Result.ToString)

      If Result <> OperationEnumerations.Result.NotProcessed Then
        lobjIdentifierBuilder.AppendFormat(" - ProcessingTime {0}", ProcessingTime.DebuggerIdentifier())
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

  Public Overridable Function ToJsonString() As String Implements IOperableResultSummary.ToJsonString
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

          .WritePropertyName("AverageProcessingTime")
          .WriteValue(ProcessingTime.AverageProcessingTime.ToString)

          '.WritePropertyName("SampleSize")
          '.WriteValue(SampleSize.ToString)
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

  Public Overridable Function ToXmlString() As String Implements IOperableResultSummary.ToXmlString
    Try
      Return Serializer.Serialize.XmlString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overridable Function ToXmlElementString() As String Implements IOperableResultSummary.ToXmlElementString
    Try
      Return Serializer.Serialize.XmlElementString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Sub ReadResultsXml(ByRef sender As IOperableResultSummary, ByVal element As XmlElement)

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
    If Not disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
      End If

      ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' TODO: set large fields to null.
    End If
    disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    ' TODO: uncomment the following line if Finalize() is overridden above.
    ' GC.SuppressFinalize(Me)
  End Sub

  Public Function GetSchema() As XmlSchema Implements IXmlSerializable.GetSchema
    Throw New NotImplementedException()
  End Function

  Public Sub ReadXml(reader As XmlReader) Implements IXmlSerializable.ReadXml
    Try

      Dim lstrScope As String = Nothing
      Dim lstrCurrentElementName As String = String.Empty

      Dim lobjProcessingTimeInfo As ProcessingTimeInfo
      Dim lobjFileCountInfo As New FileCountInfo
      Dim lobjFileSizeInfo As New FileSizeInfo
      Dim lobjVersionCountInfo As New VersionCountInfo

      With reader
        Name = .GetAttribute("name")
        DateTime.TryParse(.GetAttribute("createDate"), CreateDate)
        Node = .GetAttribute("node")
        lstrScope = .GetAttribute("scope")
        If Not String.IsNullOrEmpty(lstrScope) Then
          Scope = CType([Enum].Parse(GetType(OperationScope), lstrScope), OperationScope)
        End If
        Result = CType([Enum].Parse(GetType(Result), .GetAttribute("result")), OperationEnumerations.Result)

        .Read()
        lobjProcessingTimeInfo = New ProcessingTimeInfo()
        lobjProcessingTimeInfo.ReadXml(reader)
        ProcessingTime = lobjProcessingTimeInfo

        If TypeOf Me Is IProcessResultSummary Then
          Dim lobjOperationResultSummary As OperationResultSummary
          Do Until reader.NodeType = XmlNodeType.EndElement AndAlso (reader.Name.EndsWith("OperationResults")) _
            'OrElse reader.Name = "ProcessResultsSummary")
            If reader.NodeType = XmlNodeType.Element Then
              lstrCurrentElementName = reader.Name
              Select Case lstrCurrentElementName
                Case "OperationResultSummary"
                  lobjOperationResultSummary = New OperationResultSummary
                  lobjOperationResultSummary.ReadXml(reader)
                  DirectCast(Me, IProcessResultSummary).OperationResults.Add(lobjOperationResultSummary)
                Case "FileSizeInfo"
                  lobjFileSizeInfo.ReadXml(reader)
                  DirectCast(Me, IProcessResultSummary).FileSizeInfo = lobjFileSizeInfo
                Case "FileCountInfo"
                  lobjFileCountInfo.ReadXml(reader)
                  DirectCast(Me, IProcessResultSummary).FileCountInfo = lobjFileCountInfo
                Case "VersionCountInfo"
                  lobjVersionCountInfo.ReadXml(reader)
                  DirectCast(Me, IProcessResultSummary).VersionCountInfo = lobjVersionCountInfo
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

  Public Sub WriteXml(writer As XmlWriter) Implements IXmlSerializable.WriteXml
    Try

      With writer

        .WriteAttributeString("name", Name)
        .WriteAttributeString("createDate", CreateDate)
        .WriteAttributeString("node", Node)
        .WriteAttributeString("scope", Scope.ToString)
        .WriteAttributeString("result", Result.ToString)
        .WriteRaw(ProcessingTime.ToXmlElementString())
        'WritePreOperationEventResults(writer)

        If TypeOf Me Is IProcessResultSummary Then
          With DirectCast(Me, IProcessResultSummary)
            writer.WriteRaw(.FileSizeInfo.ToXmlElementString())
            writer.WriteRaw(.FileCountInfo.ToXmlElementString())
            writer.WriteRaw(.VersionCountInfo.ToXmlElementString())

          End With

          ' Start the OperationResults Element
          .WriteStartElement("OperationResults")
          For Each lobjOperationResultSummary As IOperableResultSummary In
            DirectCast(Me, IProcessResultSummary).OperationResults
            .WriteRaw(lobjOperationResultSummary.ToXmlElementString)
          Next
          ' End the OperationResults element
          .WriteEndElement()
        End If

      End With

      'WritePostOperationEventResults(writer)
      'WriteOnFailureEventResults(writer)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region
End Class
