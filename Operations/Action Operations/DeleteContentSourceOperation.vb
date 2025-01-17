'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  DeleteContentSourceOperation.vb
'   Description :  [type_description_here]
'   Created     :  2/21/2013 8:58:17 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Configuration
Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class DeleteContentSourceOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "DeleteContentSource"
  Friend Const PARAM_CONTENT_SOURCE_NAME As String = "ContentSourceName"
  Friend Const DEFAULT_CONTENT_SOURCE_NAME As String = "ContentSource Name"

#End Region

#Region "Class Variables"

  Private mstrContentSourceName As String = String.Empty

#End Region

#Region "Public Overrides Methods"

  Public Overrides ReadOnly Property CanRollback As Boolean
    Get
      Try
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

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

  Friend Overrides Function OnExecute() As Result
    Try
      Return DeleteContentSource()
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

      If lobjParameters.Contains(PARAM_CONTENT_SOURCE_NAME) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_CONTENT_SOURCE_NAME, DEFAULT_CONTENT_SOURCE_NAME,
          "The name of the content source to be deleted."))
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

  Private Function DeleteContentSource() As OperationEnumerations.Result
    Try
      mstrContentSourceName = GetStringParameterValue(PARAM_CONTENT_SOURCE_NAME, DEFAULT_CONTENT_SOURCE_NAME)

      If String.IsNullOrEmpty(mstrContentSourceName) Then
        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to create content source without a valid content source name."))
        Return menuResult
      End If

      If String.Equals(mstrContentSourceName, DEFAULT_CONTENT_SOURCE_NAME) Then
        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to create content source with the default content source name."))
        Return menuResult
      End If

      If ConnectionSettings.Instance.ContentSourceNames.Contains(mstrContentSourceName) Then
        Dim lstrConnectionString As String = ConnectionSettings.Instance.GetConnectionString(mstrContentSourceName)
        ConnectionSettings.Instance.ContentSourceConnectionStrings.Remove(lstrConnectionString)
        ConnectionSettings.Instance.Save()

        menuResult = OperationEnumerations.Result.Success
        Me.ProcessedMessage = String.Format("Deleted content source '{0}'", mstrContentSourceName)

      Else
        menuResult = OperationEnumerations.Result.Failed
        Me.ProcessedMessage = String.Format("Content Source '{0}' does not exist", mstrContentSourceName)
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, ex))
    End Try

    Return menuResult

  End Function

#End Region

End Class
