' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  CancelCheckoutOperation.vb
'  Description :  [type_description_here]
'  Created     :  12/15/2011 8:18:16 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Public Class CancelCheckoutOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "CancelCheckout"
  Private Shadows ReadOnly OPERATION_DESCRIPTION As String = "Cancels any current checkout state for the current item.."

  Private Const PARAM_FAIL_IF_NOT_CHECKED_OUT As String = "FailIfNotCheckedOut"

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

  Public Overrides ReadOnly Property Description As String
    Get
      Try
        Return OPERATION_DESCRIPTION
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

  Public Property FailIfNotCheckedOut As Boolean
    Get
      Try
        Return CBool(GetParameterValue(PARAM_FAIL_IF_NOT_CHECKED_OUT, False))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As Boolean)
      Try
        SetParameterValue(PARAM_FAIL_IF_NOT_CHECKED_OUT, value)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

#End Region

#Region "Friend Methods"

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try

      Dim lenuResult As Result

      lenuResult = CancelCheckout()

      If lenuResult = OperationEnumerations.Result.Failed Then
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to cancel checkout."))
      End If

      Return lenuResult

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

      If lobjParameters.Contains(PARAM_FAIL_IF_NOT_CHECKED_OUT) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_FAIL_IF_NOT_CHECKED_OUT, False,
          "Specifies whether or not the operation should fail if the document is not checked out."))
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

  Private Function CancelCheckout() As Result

    Dim lobjBCSProvider As IBasicContentServicesProvider

    Try

      RunPreOperationChecks(False)

      lobjBCSProvider = CType(PrimaryConnection.Provider.GetInterface(ProviderClass.BasicContentServices), IBasicContentServicesProvider)

      If (lobjBCSProvider IsNot Nothing) Then
        Dim lblnIsCheckedOut As Boolean = lobjBCSProvider.IsCheckedOut(Me.DocumentId)

        menuResult = ConvertResult(lobjBCSProvider.CancelCheckoutDocument(Me.DocumentId))

        If FailIfNotCheckedOut = False AndAlso lblnIsCheckedOut = False Then
          Me.ProcessedMessage = String.Format("Document {0} was not checked out.", Me.DocumentId)
          If lblnIsCheckedOut = False AndAlso menuResult = OperationEnumerations.Result.Failed Then
            menuResult = OperationEnumerations.Result.Success
          End If
        End If

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
