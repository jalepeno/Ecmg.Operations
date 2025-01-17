' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  DeleteOperation.vb
'  Description :  [type_description_here]
'  Created     :  11/29/2011 11:11:21 AM
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

Public Class DeleteOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "Delete"
  Friend Const PARAM_DELETE_ALL_VERSIONS As String = "DeleteAllVersions"

#End Region

#Region "Public Overrides Methods"

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try

      menuResult = DeleteDocument()
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
      Return False
    End Get
  End Property

#End Region

#Region "Protected Methods"

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = New Parameters

      If lobjParameters.Contains(PARAM_DELETE_ALL_VERSIONS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_DELETE_ALL_VERSIONS, True,
          "Specifies whether or not to delete all versions of the document."))
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

  Private Function DeleteDocument() As Result

    Try

      Dim lobjDelete As IDelete
      Dim lstrDocId As String = Nothing
      Dim lblnDeleteAllVersions As Boolean = GetBooleanParameterValue(PARAM_DELETE_ALL_VERSIONS, True)

      If Me.Scope = OperationScope.Source Then
        lobjDelete = CType(SourceConnection.Provider, IDelete)
        lstrDocId = Me.WorkItem.SourceDocId
      Else
        lobjDelete = CType(DestinationConnection.Provider, IDelete)
        lstrDocId = Me.WorkItem.DestinationDocId
      End If

      If String.IsNullOrEmpty(lstrDocId) Then
        Throw New InvalidOperationException(String.Format("The {0} document id is not set.", Me.Scope.ToString))
      End If

      If (lobjDelete IsNot Nothing) Then
        menuResult = ConvertResult(lobjDelete.DeleteDocument(Me.WorkItem.SourceDocId, String.Empty, lblnDeleteAllVersions))
      Else
        menuResult = OperationEnumerations.Result.Failed
      End If

    Catch ex As Exception
      If ex.Message.Contains("could not be found") Then
        ' Strip out the file name from this error message to allow these errors to be aggregated.
        Me.ProcessedMessage = String.Format("Delete Failed: {0}", "Item could not be found")
      Else
        Me.ProcessedMessage = String.Format("Delete Failed: {0}", ex.Message)
      End If
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, ex))
    End Try

    Return menuResult

  End Function

#End Region

End Class