' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ProcessTestOperation.vb
'  Description :  [type_description_here]
'  Created     :  01/23/2011 8:53:27 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------


#Region "Imports"

Imports System.Threading
Imports Documents.Utilities

#End Region

Public Class ProcessTestOperation
  Inherits Operation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "ProcessTest"

#End Region

#Region "Public Properties"

  Public Overrides ReadOnly Property Name As String
    Get
      Try
        Return OPERATION_NAME
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Overrides ReadOnly Property CanRollback As Boolean
    Get
      Return False
    End Get
  End Property

#End Region

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try
      ' <Modified by: Ernie at 1/9/2013-3:44:33 PM on machine: ERNIE-THINK>
      ' Modified for process bleed testing
      'Me.ProcessedMessage = String.Format("ItemId: {0}; ParentId: {1}; SourceDocId: {2}", Me.WorkItem.Id, Me.Parent.Id, Me.WorkItem.SourceDocId)
      'Return OperationEnumerations.Result.Success

      Static lstrBatchId As String = String.Empty
      Static lstrPreviousBatchId As String = String.Empty
      Static lstrProcessId As String = String.Empty
      Static lstrPreviousProcessId As String = String.Empty

      Me.ProcessedMessage = String.Format("ItemId: {0}; BatchId: {1}; ProcessId: {2}",
                                          Me.WorkItem.Id, Me.Parent.Id, Me.InstanceId)

      If String.IsNullOrEmpty(lstrBatchId) Then
        lstrBatchId = Me.Parent.Id
      ElseIf Me.Parent.Id <> lstrBatchId Then
        lstrPreviousBatchId = lstrBatchId
        lstrBatchId = Me.Parent.Id
      End If

      If String.IsNullOrEmpty(lstrProcessId) Then
        lstrProcessId = Me.InstanceId
      ElseIf Me.InstanceId <> lstrProcessId Then
        lstrPreviousProcessId = lstrProcessId
        menuResult = OperationEnumerations.Result.Failed
        lstrProcessId = Me.InstanceId
      End If

      Thread.Sleep(New Random().Next(300, 5000))

      menuResult = OperationEnumerations.Result.Success

      'Debug.Print(String.Format("BatchID: {0}, ProcessInstance:{1}", Me.Parent.Id, Me.InstanceId))

      ' </Modified by: Ernie at 1/9/2013-3:44:33 PM on machine: ERNIE-THINK>

      Return menuResult

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

End Class
