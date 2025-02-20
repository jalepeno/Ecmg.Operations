' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ProcessJsonConverter.vb
'  Description :  [type_description_here]
'  Created     :  01/22/2024 11:35:13 PM
'  <copyright company="Conteage">
'      Copyright (c) Conteage Corp. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Globalization
Imports Documents.Core
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Public Class ProcessJsonConverter
  Inherits JsonConverter

  Public Overrides Sub WriteJson(writer As JsonWriter, value As Object, serializer As JsonSerializer)
    Try
      Dim lobjProcess As Process = DirectCast(value, Process)

      With writer

        If Helper.IsRunningInstalled Then
          .Formatting = Formatting.None
        Else
          .Formatting = Formatting.Indented
        End If

        .WriteStartObject()

        ' Write the Operation Type
        .WritePropertyName("process")
        .WriteStartObject()

        ' Write the 'Name' property
        .WritePropertyName("name")
        .WriteValue(lobjProcess.Name)

        ' Write the 'Description' property
        .WritePropertyName("description")
        .WriteValue(lobjProcess.Description)

        ' Write the 'LogResult' property
        .WritePropertyName("logresult")
        .WriteValue(lobjProcess.LogResult)

        ' Write the 'Locale' property
        .WritePropertyName("locale")
        .WriteValue(lobjProcess.Locale.ToString())

        ' Write the 'Parameters' collection
        .WritePropertyName("parameters")
        .WriteStartArray()
        For Each lobjParameter As Parameter In lobjProcess.Parameters
          .WriteRawValue(lobjParameter.ToJson())
        Next
        .WriteEndArray()

        ' Write the 'Operations' collection
        .WritePropertyName("operations")
        .WriteStartArray()
        For Each lobjOperation As IOperation In lobjProcess.Operations
          .WriteRawValue(lobjOperation.ToJson())
        Next
        .WriteEndArray()

        If lobjProcess.RunBeforeBegin IsNot Nothing Then
          .WritePropertyName("runBeforeBegin")
          ' Write the operable element
          .WriteRaw(lobjProcess.RunBeforeBegin.ToJson())
        End If

        If lobjProcess.RunAfterComplete IsNot Nothing Then
          .WritePropertyName("runAfterComplete")
          ' Write the operable element
          .WriteStartArray()
          .WriteRaw(lobjProcess.RunAfterComplete.ToJson())
          .WriteEndArray()
        End If

        If lobjProcess.RunOnFailure IsNot Nothing Then
          .WritePropertyName("runOnFailure")
          ' Write the operable element
          .WriteStartArray()
          .WriteRaw(lobjProcess.RunOnFailure.ToJson())
          .WriteEndArray()
        End If

        If lobjProcess.RunBeforeParentBegin IsNot Nothing Then
          .WritePropertyName("runBeforeParentBegin")
          ' Write the operable element
          .WriteStartArray()
          .WriteRaw(lobjProcess.RunBeforeParentBegin.ToJson())
          .WriteEndArray()
        End If

        If lobjProcess.RunAfterParentComplete IsNot Nothing Then
          .WritePropertyName("runAfterParentComplete")
          ' Write the operable element
          .WriteStartArray()
          .WriteRaw(lobjProcess.RunAfterParentComplete.ToJson())
          .WriteEndArray()
        End If

        If lobjProcess.RunOnParentFailure IsNot Nothing Then
          .WritePropertyName("runOnParentFailure")
          ' Write the operable element
          .WriteStartArray()
          .WriteRaw(lobjProcess.RunOnParentFailure.ToJson())
          .WriteEndArray()
        End If

        If lobjProcess.RunBeforeJobBegin IsNot Nothing Then
          .WritePropertyName("runBeforeJobBegin")
          ' Write the operable element
          .WriteStartArray()
          .WriteRaw(lobjProcess.RunBeforeJobBegin.ToJson())
          .WriteEndArray()
        End If

        If lobjProcess.RunAfterJobComplete IsNot Nothing Then
          .WritePropertyName("runAfterJobComplete")
          ' Write the operable element
          .WriteStartArray()
          .WriteRaw(lobjProcess.RunAfterJobComplete.ToJson())
          .WriteEndArray()
        End If

        If lobjProcess.RunOnJobFailure IsNot Nothing Then
          .WritePropertyName("runOnJobFailure")
          ' Write the operable element
          .WriteStartArray()
          .WriteRaw(lobjProcess.RunOnJobFailure.ToJson())
          .WriteEndArray()
        End If


        .WriteEndObject()
        .WriteEndObject()

      End With


    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overrides Function ReadJson(reader As JsonReader, objectType As Type, existingValue As Object, serializer As JsonSerializer) As Object
    Try

      Dim lstrCurrentPropertyName As String = String.Empty
      Dim lstrName As String = String.Empty
      Dim lstrDescription As String = String.Empty
      Dim lblnLogResult As Boolean = False
      Dim lstrLocale As String = String.Empty
      Dim lobjOperations As New Operations
      Dim lobjParameters As New Parameters
      Dim lobjProcess As IProcess

      Dim lobjRunAfterComplete As IOperable

      While reader.Read
        Select Case reader.TokenType
          Case JsonToken.PropertyName
            lstrCurrentPropertyName = reader.Value

          Case JsonToken.String, JsonToken.Boolean, JsonToken.Date, JsonToken.Integer, JsonToken.Float
            Select Case lstrCurrentPropertyName
              Case "name"
                lstrName = reader.Value
              Case "description"
                lstrDescription = reader.Value
              Case "logresult"
                lblnLogResult = reader.Value
              Case "locale"
                lstrLocale = reader.Value
              Case "parameters"
                lstrCurrentPropertyName = reader.Value
              Case "operations"
                lstrCurrentPropertyName = reader.Value
              Case "runBeforeBegin"
                lstrCurrentPropertyName = reader.Value
              Case "runAfterComplete"
                lstrCurrentPropertyName = reader.Value
              Case "runOnFailure"
                lstrCurrentPropertyName = reader.Value
              Case "runBeforeParentBegin"
                lstrCurrentPropertyName = reader.Value
              Case "runAfterParentComplete"
                lstrCurrentPropertyName = reader.Value
              Case "runOnParentFailure"
                lstrCurrentPropertyName = reader.Value
              Case "runBeforeJobBegin"
                lstrCurrentPropertyName = reader.Value
              Case "runAfterJobComplete"
                lstrCurrentPropertyName = reader.Value
              Case "runOnJobFailure"
                lstrCurrentPropertyName = reader.Value
            End Select

          Case JsonToken.StartObject
            Select Case lstrCurrentPropertyName
              Case "parameters"
                lobjParameters.Add(Parameter.CreateFromJsonReader(reader))
              Case "operations"
                Dim lobjOperation As IOperation = Operation.CreateFromJsonReader(reader)
                If lobjOperation IsNot Nothing Then
                  lobjOperations.Add(lobjOperation)
                Else
                  Throw New InvalidOperationException()
                End If
              Case "runAfterComplete"
                lobjRunAfterComplete = Operation.CreateFromJsonReader(reader)
            End Select

        End Select
      End While

      lobjProcess = New Process(lstrName, lstrDescription, CultureInfo.CreateSpecificCulture(lstrLocale))
      With lobjProcess
        .LogResult = lblnLogResult
        .Parameters.AddRange(lobjParameters)
        .Operations.AddRange(lobjOperations)
        If lobjRunAfterComplete IsNot Nothing Then
          .RunAfterComplete = lobjRunAfterComplete
        End If
      End With

      Return lobjProcess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Function CanConvert(objectType As Type) As Boolean
    Return objectType = GetType(Process)
  End Function

End Class
