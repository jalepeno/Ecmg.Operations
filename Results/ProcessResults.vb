' ********************************************************************************
' '  Document    :  ProcessResults.vb
' '  Description :  [type_description_here]
' '  Created     :  11/13/2012-16:21:20
' '  <copyright company="ECMG">
' '      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

Imports System.IO
Imports System.Text
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Public Class ProcessResults
  Inherits CCollection(Of ProcessResult)
  Implements IProcessResults
  Implements IDisposable

#Region "Class Variables"

  Private mobjEnumerator As IEnumeratorConverter(Of IProcessResult)

#End Region

#Region "IProcessResults Implementation"

  'Public Overloads Property Item(name As String) As IProcessResult Implements IProcessResults.Item
  '  Get
  '    Try
  '      Return MyBase.Item(name)
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  '  Set(ByVal value As IProcessResult)
  '    Try
  '      MyBase.Item(name) = CType(value, ProcessResult)
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Set
  'End Property

  Public Shadows Function GetItemByName(name As String) As IProcessResult Implements IProcessResults.GetItemByName
    Try
      Return MyBase.GetItemByName(name)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shadows Sub SetItemByName(name As String, value As IProcessResult) Implements IProcessResults.SetItemByName
    Try
      MyBase.SetItemByName(name, CType(value, ProcessResult))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Function GetOperationResultsByName(lpName As String) As IOperableResults Implements IProcessResults.GetOperationResultsByName
    Try
      'Return From result In Me Where Group By Name = result.Name Into ResultsByName = Group, Count()
      Dim lobjOperableResults As New OperationResults

      For Each lobjProcessResult As ProcessResult In Me
        lobjOperableResults.Add(lobjProcessResult.OperationResults.GetItemByName(lpName))
      Next

      Return lobjOperableResults

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetOperationResultsByIndex(index As Integer) As IOperableResults Implements IProcessResults.GetOperationResultsByIndex
    Try
      'Return From result In Me Group By Name = result.Name Into ResultsByName = Group, Count()
      Throw New NotImplementedException
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function ToJsonString() As String Implements IProcessResults.ToJsonString
    Try
      ' We will do manual JSON serialization for complete speed and control

      Dim lobjStringBuilder As New StringBuilder
      Dim lobjStringWriter As New StringWriter(lobjStringBuilder)

      Using lobjJSONWriter As New JsonTextWriter(lobjStringWriter)
        With lobjJSONWriter
          .Formatting = Newtonsoft.Json.Formatting.Indented
          .WriteStartObject()
          .WritePropertyName("ProcessResults")
          .WriteStartArray()
          For Each lobjResult As IProcessResult In Items
            .WriteRawValue(lobjResult.ToJsonString)
          Next
          .WriteEndArray()
          .WriteEndObject()
        End With
      End Using

      Return lobjStringBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function ToXmlString() As String Implements IProcessResults.ToXmlString
    Try
      Return Serializer.Serialize.XmlString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function ToXmlElementString() As String Implements IProcessResults.ToXmlElementString
    Try
      Return Serializer.Serialize.XmlElementString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Sub Add(item As IProcessResult) Implements ICollection(Of IProcessResult).Add
    Try
      Add(CType(item, ProcessResult))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Sub Add(item As IProcessResults)
    Try
      For Each lobjProcessResult As IProcessResult In item
        Add(lobjProcessResult)
      Next
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Sub Clear() Implements ICollection(Of IProcessResult).Clear
    Try
      MyBase.Clear()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Function Contains(item As IProcessResult) As Boolean Implements ICollection(Of IProcessResult).Contains
    Try
      Return Contains(CType(item, ProcessResult))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Sub CopyTo(array() As IProcessResult, arrayIndex As Integer) Implements ICollection(Of IProcessResult).CopyTo
    Try
      MyBase.CopyTo(CType(array, ProcessResult()), arrayIndex)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads ReadOnly Property Count As Integer Implements ICollection(Of IProcessResult).Count
    Get
      Try
        Return MyBase.Count
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Overloads ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of IProcessResult).IsReadOnly
    Get
      Try
        Return MyBase.IsReadOnly
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Overloads Function Remove(item As IProcessResult) As Boolean Implements ICollection(Of IProcessResult).Remove
    Try
      Return Remove(CType(item, ProcessResult))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function GetEnumerator() As IEnumerator(Of IProcessResult) Implements IEnumerable(Of IProcessResult).GetEnumerator
    Try
      Return IPropertyEnumerator.GetEnumerator
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Properties"

  Protected ReadOnly Property IsDisposed() As Boolean
    Get
      Return disposedValue
    End Get
  End Property

  Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IProcessResult)
    Get
      Try
        If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
          mobjEnumerator = New IEnumeratorConverter(Of IProcessResult)(Me.ToArray, GetType(ProcessResult), GetType(IProcessResult))
        End If
        Return mobjEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

#End Region

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        ' DISPOSETODO: dispose managed state (managed objects).
        MyBase.Dispose()
      End If

      ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' DISPOSETODO: set large fields to null.
    End If
    Me.disposedValue = True
  End Sub

#End Region

End Class
