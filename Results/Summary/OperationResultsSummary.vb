' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IOperableResultsSummary.vb
'  Description :  [type_description_here]
'  Created     :  4/26/2016 10:50:40 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml
Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class OperationResultsSummary
  Inherits CCollection(Of OperationResultSummary)
  Implements IOperableResultsSummary
  Implements IDisposable

#Region "Class Variables"

  Private mobjEnumerator As IEnumeratorConverter(Of IOperableResultSummary)

#End Region

#Region "Constructors"

  Public Sub New()
    MyBase.New()
  End Sub

  Public Sub New(ByVal lpParentXMLElement As XmlNode)
    MyBase.new()
    Try
      InitializeFromXml(lpParentXMLElement)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Sub

#End Region

#Region "IOperableResultsSummary Implementation"

  Public Shadows Function GetItemByName(name As String) As IOperableResultSummary Implements IOperableResultsSummary.GetItemByName
    Try
      Return MyBase.GetItemByName(name)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shadows Sub SetItemByName(name As String, value As IOperableResultSummary) Implements IOperableResultsSummary.SetItemByName
    Try
      MyBase.SetItemByName(name, CType(value, OperationResultSummary))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads ReadOnly Property Count As Integer Implements ICollection(Of IOperableResultSummary).Count
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

  Public Overloads ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of IOperableResultSummary).IsReadOnly
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

  Public Overloads Sub Add(item As IOperableResultSummary) Implements ICollection(Of IOperableResultSummary).Add
    Try
      Add(CType(item, OperationResultSummary))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Sub CopyTo(array() As IOperableResultSummary, arrayIndex As Integer) Implements ICollection(Of IOperableResultSummary).CopyTo
    Try
      MyBase.CopyTo(CType(array, OperationResultSummary()), arrayIndex)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub


  Public Overloads Sub Clear() Implements ICollection(Of IOperableResultSummary).Clear
    Try
      MyBase.Clear()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Function Contains(item As IOperableResultSummary) As Boolean Implements ICollection(Of IOperableResultSummary).Contains
    Try
      Return Contains(CType(item, OperationResultSummary))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function Remove(item As IOperableResultSummary) As Boolean Implements ICollection(Of IOperableResultSummary).Remove
    Try
      Return Remove(CType(item, OperationResultSummary))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of IOperableResultSummary) Implements System.Collections.Generic.IEnumerable(Of IOperableResultSummary).GetEnumerator
    Try
      Return IPropertyEnumerator.GetEnumerator
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function ToXmlElementString() As String Implements IOperableResultsSummary.ToXmlElementString
    Throw New NotImplementedException()
  End Function

  'Private Function IEnumerable_GetEnumerator() As IEnumerator(Of IOperableResultSummary) Implements IEnumerable(Of IOperableResultSummary).GetEnumerator
  '  Throw New NotImplementedException()
  'End Function

#End Region


#Region "Private Properties"

  Protected ReadOnly Property IsDisposed() As Boolean
    Get
      Return disposedValue
    End Get
  End Property

  Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IOperableResultSummary)
    Get
      Try
        If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
          mobjEnumerator = New IEnumeratorConverter(Of IOperableResultSummary)(Me.ToArray, GetType(OperationResultSummary), GetType(IOperableResultSummary))
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

#Region "Private Methods"

  Private Sub InitializeFromXml(ByVal lpParentXMLElement As XmlNode)
    Try
      If lpParentXMLElement IsNot Nothing AndAlso lpParentXMLElement.HasChildNodes Then
        For Each lobjResultNode As XmlElement In lpParentXMLElement.ChildNodes
          'If Not String.IsNullOrEmpty(lobjResultNode.InnerXml) AndAlso lobjResultNode.InnerXml <> "<ChildOperations />" Then
          Add(New OperationResultSummary(lobjResultNode))
          'End If
        Next
      End If
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
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        ' DISPOSETODO: dispose managed state (managed objects).

        For Each lobjResult As OperationResultSummary In Me
          lobjResult.Dispose
        Next
        Me.Clear()

      End If

      ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' DISPOSETODO: set large fields to null.
    End If
    Me.disposedValue = True
  End Sub

#End Region
End Class
