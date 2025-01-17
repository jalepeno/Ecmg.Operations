' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperableEventArgs.vb
'  Description :  [type_description_here]
'  Created     :  11/18/2011 4:35:00 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

''' <summary>
''' Base class for operable events.
''' </summary>
''' <remarks></remarks>
<DebuggerDisplay("{DebuggerIdentifier(),nq}")>
Public Class OperableEventArgs
  Inherits EventArgs

#Region "Class Variables"

  Private mobjOperation As IOperable = Nothing
  Private mobjWorkItem As IWorkItem = Nothing

#End Region

#Region "Public Properties"

  Public ReadOnly Property WorkItem As IWorkItem
    Get
      Return mobjWorkItem
    End Get
  End Property

  Public ReadOnly Property Operation As IOperable
    Get
      Return mobjOperation
    End Get
  End Property

  Public ReadOnly Property DocumentId() As String
    Get
      Return mobjOperation.DocumentId
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New()

  End Sub

  Public Sub New(lpOperable As IOperable)
    Me.New(lpOperable, lpOperable.WorkItem)
  End Sub

  Public Sub New(lpOperable As IOperable, lpWorkItem As IWorkItem)
    Try
      mobjOperation = lpOperable
      mobjWorkItem = lpWorkItem
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Protected Methods"

  Protected Friend Overridable Function DebuggerIdentifier() As String
    Try
      If Not String.IsNullOrEmpty(DocumentId) Then
        Return String.Format("DocumentId={0}", DocumentId)
      Else
        Return "OperationEventArgs"
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class