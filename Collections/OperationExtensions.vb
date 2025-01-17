'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  OperationExtensions.vb
'   Description :  [type_description_here]
'   Created     :  1/9/2015 11:39:53 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities
Imports Operations.Extensions

#End Region

Public Class OperationExtensions
  Inherits CCollection(Of OperationExtension)
  Implements IOperationExtensions

#Region "Class Variables"

  Private mobjEnumerator As IEnumeratorConverter(Of IOperationExtension)

#End Region

#Region "IOperationExtensions Implementation"

  Public Overloads Sub Add(item As Extensions.IOperationExtension) Implements ICollection(Of Extensions.IOperationExtension).Add
    Try
      Add(CType(item, OperationExtension))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Sub Add(item As OperationExtension)
    Try
      ' MyBase.Add(item)
      If String.IsNullOrEmpty(item.Name()) Then
        Exit Sub
      End If
      If Contains(item.Name) = False Then
        MyBase.Add(item)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Sub Clear() Implements ICollection(Of Extensions.IOperationExtension).Clear
    Try
      MyBase.Clear()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Function Contains(item As Extensions.IOperationExtension) As Boolean Implements ICollection(Of Extensions.IOperationExtension).Contains
    Try
      Return Contains(CType(item, OperationExtension))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Sub CopyTo(array() As Extensions.IOperationExtension, arrayIndex As Integer) Implements ICollection(Of Extensions.IOperationExtension).CopyTo
    Try
      MyBase.CopyTo(array, arrayIndex)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads ReadOnly Property Count As Integer Implements ICollection(Of Extensions.IOperationExtension).Count
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

  Public Overloads ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of Extensions.IOperationExtension).IsReadOnly
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

  Public Overloads Function Remove(item As Extensions.IOperationExtension) As Boolean Implements ICollection(Of Extensions.IOperationExtension).Remove
    Try
      Return Remove(CType(item, OperationExtension))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function GetEnumerator() As IEnumerator(Of Extensions.IOperationExtension) Implements IEnumerable(Of Extensions.IOperationExtension).GetEnumerator
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

  Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IOperationExtension)
    Get
      Try
        If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
          mobjEnumerator = New IEnumeratorConverter(Of IOperationExtension)(Me.ToArray, GetType(OperationExtension), GetType(IOperationExtension))
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

End Class
