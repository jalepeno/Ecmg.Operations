' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  CheckinOperation.vb
'  Description :  [type_description_here]
'  Created     :  12/12/2011 4:50:16 PM
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

Public Class CheckinOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "Checkin"
  Private Const PARAM_CHECKIN_AS_MAJOR As String = "CheckinAsMajor"

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

  Public Property CheckinAsMajor As Boolean
    Get
      Try
        Return GetBooleanParameterValue(PARAM_CHECKIN_AS_MAJOR, False)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As Boolean)
      Try
        SetParameterValue(PARAM_CHECKIN_AS_MAJOR, value)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
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

      Return CheckInDocument()

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

      If lobjParameters.Contains(PARAM_CHECKIN_AS_MAJOR) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_CHECKIN_AS_MAJOR, False,
          "Specifies whether to check in the document as a major or minor version."))
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

  Private Function CheckInDocument() As Result

    Dim lobjBCSProvider As IBasicContentServicesProvider

    Try

      RunPreOperationChecks(False)

      lobjBCSProvider = CType(PrimaryConnection.Provider.GetInterface(ProviderClass.BasicContentServices), IBasicContentServicesProvider)

      If (lobjBCSProvider IsNot Nothing) Then

        ' Get the content container to check in
        Dim lobjLatestContent As Content = Me.WorkItem.Document.LatestVersion.PrimaryContent
        Dim lobjContentContainer As IContentContainer = New ContentStreamContainer(lobjLatestContent)

        menuResult = ConvertResult(lobjBCSProvider.CheckinDocument(Me.DocumentId, lobjContentContainer, CheckinAsMajor))

      Else
        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to get basic content services interface"))
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
