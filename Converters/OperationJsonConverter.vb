' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperationJsonConverter.vb
'  Description :  [type_description_here]
'  Created     :  01/22/2024 2:12:13 PM
'  <copyright company="Conteage">
'      Copyright (c) Conteage Corp. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Public Class OperationJsonConverter
  Inherits JsonConverter

  Public Overrides Sub WriteJson(writer As JsonWriter, value As Object, serializer As JsonSerializer)
    Try
      Dim lobjOperation As IOperation = DirectCast(value, IOperation)

      With writer
        If Helper.IsRunningInstalled Then
          .Formatting = Formatting.None
        Else
          .Formatting = Formatting.Indented
        End If

        .WriteStartObject()

        ' Write the Operation Type
        .WritePropertyName("type")
        .WriteValue(lobjOperation.GetType.Name)

        ' Write the 'Name' property
        .WritePropertyName("name")
        .WriteValue(lobjOperation.Name)

        ' Write the 'Description' property
        .WritePropertyName("description")
        .WriteValue(lobjOperation.Description)

        ' Write the 'LogResult' property
        .WritePropertyName("logresult")
        .WriteValue(lobjOperation.LogResult)

        ' Write the 'Scope' property
        .WritePropertyName("scope")
        .WriteValue([Enum].GetName(GetType(OperationScope), lobjOperation.Scope))

        ' Write the 'Parameters' collection
        .WritePropertyName("parameters")
        .WriteStartArray()
        For Each lobjParameter As Parameter In lobjOperation.Parameters
          .WriteRawValue(lobjParameter.ToJson())
        Next

        .WriteEndArray()

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
      Dim lstrType As String = String.Empty
      Dim lstrName As String = String.Empty
      Dim lstrDescription As String = String.Empty
      Dim lblnLogResult As Boolean = False
      Dim lstrScope As String = String.Empty
      Dim lobjParameters As New Parameters

      While reader.Read
        Select Case reader.TokenType
          Case JsonToken.PropertyName
            lstrCurrentPropertyName = reader.Value

          Case JsonToken.String, JsonToken.Boolean, JsonToken.Date, JsonToken.Integer, JsonToken.Float
            Select Case lstrCurrentPropertyName
              Case "type"
                lstrType = reader.Value
              Case "name"
                lstrName = reader.Value
              Case "description"
                lstrDescription = reader.Value
              Case "logresult"
                lblnLogResult = reader.Value
              Case "scope"
                lstrScope = reader.Value
              Case "parameters"
                lstrCurrentPropertyName = reader.Value
                ' TODO: Handle the parameters

            End Select

          Case JsonToken.StartArray
            Select Case lstrCurrentPropertyName
              Case "parameters"
                'Beep()
                ' TODO: Handle the parameters
                lobjParameters.Add(Parameter.CreateFromJsonReader(reader))
            End Select

          Case JsonToken.StartObject
            Select Case lstrCurrentPropertyName
              Case "parameters"
                'Beep()
                ' TODO: Handle the parameters
                lobjParameters.Add(Parameter.CreateFromJsonReader(reader))
            End Select
          Case JsonToken.EndArray
            Dim lstrOperationName As String = lstrType.Replace("Operation", String.Empty)
            If OperationFactory.Instance.AvailableOperations.ContainsKey(lstrOperationName) Then
              Dim lobjOperationType As Type = OperationFactory.Instance.AvailableOperations.Item(lstrOperationName)
              Dim lobjOperation As IOperation = OperationFactory.Create(lstrOperationName)
              With lobjOperation
                .LogResult = lblnLogResult
                .Scope = [Enum].Parse(GetType(OperationScope), lstrScope)
                .Parameters = lobjParameters
              End With
              Return lobjOperation
            Else
              Throw New UnknownOperationException(lstrOperationName)
            End If
        End Select
      End While

      Return Nothing

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Function CanConvert(objectType As Type) As Boolean
    If TypeOf objectType Is IOperation OrElse Helper.IsAssignableFrom("IOperation", objectType) Then
      Return True
    Else
      Return False
    End If
    'Return objectType = GetType(IOperation)
  End Function

End Class
