﻿'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  MigrateFolderOperation.vb
'   Description :  [type_description_here]
'   Created     :  6/9/2015 12:49:08 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Public Class MigrateFolderOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "MigrateFolder"

#End Region

#Region "Public Overrides Methods"

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try

      ' TODO: Implement operation or call implementing method here.
      ' See ExportDocument for an example

      ProcessedMessage = "Operation not yet implemented"
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, Me.ProcessedMessage))
      Return menuResult

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

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
      Return True
    End Get
  End Property

#End Region

#Region "Private Methods"

#End Region

End Class
