' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Operations.vb
'  Description :  [type_description_here]
'  Created     :  11/23/2011 4:37:32 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class Operations
  Inherits CCollection(Of IOperable)
  Implements IOperations
  Implements ICloneable

#Region "IOperations Implementation"

  'Public Event Complete(ByVal sender As Object, ByVal e As OperableEventArgs) Implements IOperable.Complete

  'Public Event Begin(ByVal sender As Object, ByVal e As OperableEventArgs) Implements IOperable.Begin

  'Public Event OperatingError(ByVal sender As Object, ByVal e As OperableErrorEventArgs) Implements IOperable.OperatingError


  'Protected Overridable Function GetBooleanParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Boolean Implements IOperable.GetBooleanParameterValue
  '  Try
  '    Throw new NotImplementedException
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  'Protected Overridable Function GetEnumParameterValue(ByVal lpParameterName As String, ByVal lpEnumType As Type, ByVal lpDefaultValue As Object) As [Enum] Implements IOperable.GetEnumParameterValue
  '  Try
  '    Throw new NotImplementedException
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  'Protected Overridable Function GetStringParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As String Implements IOperable.GetStringParameterValue
  '  Try
  '    Throw new NotImplementedException
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  'Protected Overridable Function GetParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Object Implements IOperable.GetParameterValue
  '  Try
  '    Throw New NotImplementedException
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  'Protected Overridable Function OnExecute() As OperationEnumerations.Result Implements IOperable.OnExecute
  '  Try
  '    Throw New NotImplementedException
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  Public Overrides Sub AddRange(ByVal lpOperations As IEnumerable(Of IOperable)) Implements IOperations.AddRange
    Try
      MyBase.AddRange(lpOperations)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Function Execute(ByVal lpWorkItem As IWorkItem) As OperationEnumerations.Result Implements IOperations.Execute ', IOperable.Execute
    Try
      For Each lobjOperation As IOperable In Me
        If lobjOperation.Execute(lpWorkItem) = Result.Failed Then
          Return Result.Failed
        End If
      Next

      Return Result.Success

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function Rollback(ByVal lpWorkItem As IWorkItem) As Result Implements IOperations.Rollback
    Try
      For Each lobjOperation As IOperable In Me
        If lobjOperation.CanRollback Then
          If lobjOperation.Rollback(lpWorkItem) = Result.RollbackFailed Then
            Return Result.RollbackFailed
          End If
        End If
      Next

      Return Result.RollbackSuccess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function ContainsOperation(lpOperationName As String) As Boolean Implements IOperations.ContainsOperation
    Try
      Return ContainsOperation(lpOperationName, Nothing)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function ContainsOperation(lpOperationName As String, ByRef lpFoundOperation As IOperable) As Boolean Implements IOperations.ContainsOperation
    Try
      lpFoundOperation = MyBase.FirstOrDefault(Function(Operable) Operable.Name = lpOperationName)

      If lpFoundOperation IsNot Nothing Then
        Return True
      End If

      lpFoundOperation = MyBase.FirstOrDefault(Function(Operable) Operable.DisplayName = lpOperationName)

      If lpFoundOperation IsNot Nothing Then
        Return True
      End If

      Return False
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Friend Methods"

  Friend Sub AssignHost(lpHost As Object) Implements IOperations.AssignHost
    Try
      For Each lobjOperable As IOperable In Me
        lobjOperable.Host = lpHost
      Next
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Friend Sub AssignTag(lpTag As Object) Implements IOperations.AssignTag
    Try
      For Each lobjOperable As IOperable In Me
        lobjOperable.Tag = lpTag
      Next
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

  '#Region "IOperations Implementation"

  '    Public Overloads Sub Add(item As IOperation) Implements System.Collections.Generic.ICollection(Of IOperation).Add
  '      Try
  '        If TypeOf (item) Is Operation Then
  '          MyBase.Add(CType(item, Operation))
  '        Else
  '          Throw New ArgumentException("Operation object expected", "item")
  '        End If
  '      Catch ex As Exception
  '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '        ' Re-throw the exception to the caller
  '        Throw
  '      End Try
  '    End Sub

  '    Public Overloads Sub Clear() Implements System.Collections.Generic.ICollection(Of IOperation).Clear
  '      Try
  '        MyBase.Clear()
  '      Catch ex As Exception
  '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '        ' Re-throw the exception to the caller
  '        Throw
  '      End Try
  '    End Sub

  '    Public Overloads Function Contains(item As IOperation) As Boolean Implements System.Collections.Generic.ICollection(Of IOperation).Contains
  '      Try
  '        Return MyBase.Contains(item.Name)
  '      Catch ex As Exception
  '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '        ' Re-throw the exception to the caller
  '        Throw
  '      End Try
  '    End Function

  '    Public Overloads Sub CopyTo(array() As IOperation, arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of IOperation).CopyTo
  '      Try
  '        MyBase.CopyTo(array, arrayIndex)
  '      Catch ex As Exception
  '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '        ' Re-throw the exception to the caller
  '        Throw
  '      End Try
  '    End Sub

  '    Public Overloads ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of IOperation).Count
  '      Get
  '        Try
  '          Return MyBase.Count
  '        Catch ex As Exception
  '          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '          ' Re-throw the exception to the caller
  '          Throw
  '        End Try
  '      End Get
  '    End Property

  '    Public Overloads ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of IOperation).IsReadOnly
  '      Get
  '        Try
  '          Return MyBase.IsReadOnly
  '        Catch ex As Exception
  '          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '          ' Re-throw the exception to the caller
  '          Throw
  '        End Try
  '      End Get
  '    End Property

  '    Public Overloads Function Remove(item As IOperation) As Boolean Implements System.Collections.Generic.ICollection(Of IOperation).Remove
  '      Try
  '        Return MyBase.Remove(CType(item, Operation))
  '      Catch ex As Exception
  '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '        ' Re-throw the exception to the caller
  '        Throw
  '      End Try
  '    End Function

  '    Public Overloads Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of IOperation) Implements System.Collections.Generic.IEnumerable(Of IOperation).GetEnumerator
  '      Try
  '        Return MyBase.GetEnumerator
  '      Catch ex As Exception
  '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '        ' Re-throw the exception to the caller
  '        Throw
  '      End Try
  '    End Function

  '#End Region

#Region "ICloneable Implementation"

  Public Function Clone() As Object Implements System.ICloneable.Clone
    Try
      Dim lobjOperations As New Operations

      For Each lobjOperation As IOperable In Me
        lobjOperations.Add(CType(lobjOperation.Clone, IOperable))
      Next

      Return lobjOperations

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class