' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  CheckoutOperation.vb
'  Description :  [type_description_here]
'  Created     :  12/12/2011 4:49:00 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Public Class CheckoutOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "Checkout"
  Friend Const PARAM_EXPORT_ON_CHECKOUT As String = "ExportOnCheckout"

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
      Return True
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New()
    Try

      ' Set the default scope
      Scope = OperationScope.Source

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Friend Methods"

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try

      Return CheckoutDocument()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Overrides Function OnRollback() As Result
    Try
      ' <Added by: Ernie at: 12/18/2012-4:01:48 PM on machine: ERNIE-THINK>
      ' TODO: Implement this method by cancelling the checkout
      Return OperationEnumerations.Result.RollbackNotImplemented
      ' </Added by: Ernie at: 12/18/2012-4:01:48 PM on machine: ERNIE-THINK>
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

      'If lobjParameters.Contains(PARAM_EXPORT_ON_CHECKOUT) = False Then
      '  lobjParameters.Add(New Parameter(PropertyType.ecmBoolean, PARAM_EXPORT_ON_CHECKOUT, True))
      'End If

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Function CheckoutDocument() As Result

    Dim lobjBCSProvider As IBasicContentServicesProvider

    Try

      RunPreOperationChecks(False)

      lobjBCSProvider = CType(PrimaryConnection.Provider.GetInterface(ProviderClass.BasicContentServices), IBasicContentServicesProvider)

      If (lobjBCSProvider IsNot Nothing) Then
        menuResult = ConvertResult(lobjBCSProvider.CheckoutDocument(Me.DocumentId, FileHelper.Instance.TempPath, Nothing))

        'If Parameters.Contains(PARAM_EXPORT_ON_CHECKOUT) AndAlso Parameters.Item(PARAM_EXPORT_ON_CHECKOUT).Value = True Then

        '  Dim lobjExportOperation As New ExportOperation
        '  lobjExportOperation.Parameters(ExportOperation.PARAM_GENERATE_CDF).Value = True

        '  menuResult = lobjExportOperation.Execute(Me.WorkItem)

        'End If

        If menuResult = Result.Success Then
          Me.ProcessedMessage = String.Format("Successfully checked out document '{0}'.", Me.DocumentId)
        ElseIf menuResult = Result.Failed AndAlso String.IsNullOrEmpty(Me.ProcessedMessage) Then
          Me.ProcessedMessage = String.Format("Failed to check out document '{0}'.", Me.DocumentId)
        End If

      Else
        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to get basic content services interface"))
      End If

    Catch CheckedOutEx As DocumentAlreadyCheckedOutException
      ApplicationLogging.LogException(CheckedOutEx, Reflection.MethodBase.GetCurrentMethod)
      Me.ProcessedMessage = String.Format("Failed to check out document: '{0}'.", CheckedOutEx.Message)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, CheckedOutEx))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, ex))
    End Try

    Return menuResult

  End Function

#End Region

End Class
