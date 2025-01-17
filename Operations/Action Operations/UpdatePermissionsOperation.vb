' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  UpdatePermissionsOperation.vb
'  Description :  [type_description_here]
'  Created     :  8/14/2012 12:57:40 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Security
Imports Documents.Utilities

#End Region

Public Class UpdatePermissionsOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "UpdatePermissions"
  Friend Const PARAM_SECURITY_ARGUMENTS As String = "SecurityArguments"

#End Region

#Region "Class Variables"

  Private mobjSecurityArguments As ObjectSecurityArgs = Nothing

#End Region

#Region "Public Properties"

  Public ReadOnly Property SecurityArguments As ObjectSecurityArgs
    Get
      If mobjSecurityArguments Is Nothing Then
        mobjSecurityArguments = CType(GetParameterValue(PARAM_SECURITY_ARGUMENTS, New ObjectSecurityArgs), ObjectSecurityArgs)
      End If
      Return mobjSecurityArguments
    End Get
  End Property
#End Region

#Region "Public Overrides Methods"

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

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try
      Return UpdatePermissions()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Protected Methods"

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = New Parameters

      If lobjParameters.Contains(PARAM_SECURITY_ARGUMENTS) = False Then
        Dim lobjSecurityArgs As New ObjectSecurityArgs
        lobjParameters.Add(New ObjectParameter(PropertyType.ecmObject, PARAM_SECURITY_ARGUMENTS,
          CreateSampleSecurityArgs(), "The security arguments object that specify the security to update."))
      End If

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Function UpdatePermissions() As Result
    Try

      RunPreOperationChecks(False)

      Dim lobjPermissionUpdater As IUpdatePermissions = CType(DestinationConnection.Provider.GetInterface(ProviderClass.UpdatePermissions), IUpdatePermissions)

      Dim lstrErrorMessage As String = String.Empty

      If String.IsNullOrEmpty(Me.WorkItem.SourceDocId) Then
        Throw New InvalidOperationException("No source document path available")
      End If

      'ApplicationLogging.WriteLogEntry(String.Format("Thread Id: {1}   WorkItem Id: {0}   SourceDocId: {2}  Provider Tag {3}  Process InstanceId: {4} ", Me.WorkItem.Id, Threading.Thread.CurrentThread.ManagedThreadId.ToString, Me.WorkItem.SourceDocId, Me.DestinationConnection.Provider.Tag, Me.InstanceId), TraceEventType.Information, 4577565)

      SecurityArguments.ObjectID = Me.WorkItem.SourceDocId
      SecurityArguments.Tag = Me.InstanceId

      Dim lblnSuccess As Boolean = lobjPermissionUpdater.UpdatePermissions(SecurityArguments)

      If lblnSuccess = True Then
        menuResult = OperationEnumerations.Result.Success
      Else
        menuResult = OperationEnumerations.Result.Failed
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Me.ProcessedMessage = String.Format("Update Permissions Failed: {0}", ex.Message)
      menuResult = OperationEnumerations.Result.Failed
    End Try

    Return menuResult

  End Function

  Private Function CreateSampleSecurityArgs() As ObjectSecurityArgs
    Try
      Dim lobjSecurityArgs As New DocumentSecurityArgs
      Dim lobjAccessRight As IAccessRight = New AccessLevel(PermissionLevel.FullControl)
      Dim lobjPermission As IPermission = New ItemPermission("BIAdmin", AccessType.Allow, lobjAccessRight, PermissionSource.Direct)
      lobjPermission.PrincipalType = PrincipalType.Group

      With lobjSecurityArgs
        .Mode = ObjectSecurityArgs.UpdateMode.Append
        .Permissions.Add(lobjPermission)
      End With

      Return lobjSecurityArgs

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
