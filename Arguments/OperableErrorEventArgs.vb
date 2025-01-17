' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperableErrorEventArgs.vb
'  Description :  [type_description_here]
'  Created     :  11/18/2011 4:41:48 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Public Class OperableErrorEventArgs
  Inherits OperableEventArgs

#Region "Class Variables"

  Private mobjException As Exception = Nothing
  Private mstrMessage As String = String.Empty

#End Region

#Region "Public Properties"

  Public ReadOnly Property Exception As Exception
    Get
      Try
        Return mobjException
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property Message As String
    Get
      Try
        Return mstrMessage
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New()

  End Sub

  Public Sub New(lpOperable As IOperable, lpException As Exception)
    Me.New(lpOperable, lpOperable.WorkItem, lpException)
  End Sub

  Public Sub New(lpOperable As IOperable, lpMessage As String)
    Me.New(lpOperable, lpOperable.WorkItem, lpMessage)
  End Sub

  Public Sub New(lpOperable As IOperable, lpWorkItem As IWorkItem, lpException As Exception)
    MyBase.New(lpOperable, lpWorkItem)
    Try
      mobjException = lpException
      mstrMessage = lpException.Message
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(lpOperable As IOperable, lpWorkItem As IWorkItem, lpMessage As String)
    MyBase.New(lpOperable, lpWorkItem)
    Try
      mstrMessage = lpMessage
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class