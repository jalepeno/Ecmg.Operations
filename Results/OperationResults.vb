' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperationResults.vb
'  Description :  [type_description_here]
'  Created     :  12/8/2011 8:04:37 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Public Class OperationResults
  Inherits CCollection(Of OperationResult)
  Implements IOperableResults
  Implements IDisposable

#Region "Class Variables"

  Private mobjEnumerator As IEnumeratorConverter(Of IOperableResult)

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

#Region "IOperableResults Implementation"

  'Public Overloads Property Item(ByVal name As String) As IOperableResult Implements IOperableResults.Item
  '  Get
  '    Try
  '      Return MyBase.Item(name)
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  '  Set(ByVal value As IOperableResult)
  '    Try
  '      MyBase.Item(name) = CType(value, OperationResult)
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Set
  'End Property

  Public Shadows Function GetItemByName(name As String) As IOperableResult Implements IOperableResults.GetItemByName
    Try
      Return MyBase.GetItemByName(name)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shadows Sub SetItemByName(name As String, value As IOperableResult) Implements IOperableResults.SetItemByName
    Try
      MyBase.SetItemByName(name, CType(value, OperationResult))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Sub Add(ByVal item As IOperableResult) Implements System.Collections.Generic.ICollection(Of IOperableResult).Add
    Try
      Add(CType(item, OperationResult))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Sub Clear() Implements System.Collections.Generic.ICollection(Of IOperableResult).Clear
    Try
      MyBase.Clear()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Function Contains(ByVal item As IOperableResult) As Boolean Implements System.Collections.Generic.ICollection(Of IOperableResult).Contains
    Try
      Return Contains(CType(item, OperationResult))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Sub CopyTo(ByVal array() As IOperableResult, ByVal arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of IOperableResult).CopyTo
    Try
      MyBase.CopyTo(CType(array, OperationResult()), arrayIndex)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of IOperableResult).Count
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

  Public Overloads ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of IOperableResult).IsReadOnly
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

  Public Overloads Function Remove(ByVal item As IOperableResult) As Boolean Implements System.Collections.Generic.ICollection(Of IOperableResult).Remove
    Try
      Return Remove(CType(item, OperationResult))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of IOperableResult) Implements System.Collections.Generic.IEnumerable(Of IOperableResult).GetEnumerator
    Try
      Return IPropertyEnumerator.GetEnumerator
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function ToXmlElementString() As String Implements IOperableResults.ToXmlElementString
    Try
      Return Serializer.Serialize.XmlElementString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
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

  Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IOperableResult)
    Get
      Try
        If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
          mobjEnumerator = New IEnumeratorConverter(Of IOperableResult)(Me.ToArray, GetType(OperationResult), GetType(IOperableResult))
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
          Add(New OperationResult(lobjResultNode))
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

        For Each lobjResult As OperationResult In Me
          lobjResult.Dispose()
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
