' ********************************************************************************
' '  Document    :  CreateLinkOperation.vb
' '  Description :  [type_description_here]
' '  Created     :  11/21/2012-1:32:13
' '  <copyright company="ECMG">
' '      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

Imports Documents
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Public Class CreateLinkOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "CreateLink"
  Friend Const PARAM_URL_FACTORY As String = "UrlFactory"

#End Region

#Region "Class Variables"

  Dim mobjUrlFactory As UrlFactory = Nothing

#End Region

#Region "Public Properties"

  Public Overrides ReadOnly Property CanRollback As Boolean
    Get
      Return True
    End Get
  End Property

  Public ReadOnly Property UrlFactory As UrlFactory
    Get
      Try
        Return mobjUrlFactory
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
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

  Friend Overrides Function OnExecute() As Result
    Try

      Return CreateLink()

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

      If lobjParameters.Contains(PARAM_URL_FACTORY) = False Then
        Dim lobjUrlFactoryParam As New ObjectParameter
        With lobjUrlFactoryParam
          .Name = PARAM_URL_FACTORY
          .SystemName = PARAM_URL_FACTORY
          .Description = "The factory pattern to be used for creating the link."
          .Type = PropertyType.ecmObject
          .Cardinality = Cardinality.ecmSingleValued
          .Value = New UrlFactory("Mary had a {0} whose {1} was {2}.", "little lamb", "fleece", "white as snow")
        End With
        lobjParameters.Add(lobjUrlFactoryParam)
      End If

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '   Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Function CreateLink() As Result

    Dim lobjLinkManager As Providers.ILinkManager

    Try

      RunPreOperationChecks(True)

      mobjUrlFactory = CType(GetParameterValue(PARAM_URL_FACTORY, New UrlFactory), Core.UrlFactory)

      'If Me.WorkItem.Process.Parameters.Contains("SourceJob") Then
      '  Beep()
      'End If

      If mobjUrlFactory.FormatString.Contains("{Process.SourceDocId}") AndAlso
        Me.WorkItem.Process.Parameters.Contains(Process.PARAM_SOURCE_DOC_ID) Then
        Dim lobjSourceDocIdParameter As IParameter = Me.WorkItem.Process.Parameters.Item(Process.PARAM_SOURCE_DOC_ID)
        If lobjSourceDocIdParameter.HasValue Then
          mobjUrlFactory.FormatString = mobjUrlFactory.FormatString.Replace("{Process.SourceDocId}", lobjSourceDocIdParameter.Value.ToString)
        End If
      End If

      mobjUrlFactory.ResolveParametersFromDocument(Me.WorkItem.Document)

      Dim lobjCreateDocumentLinkArgs As New Arguments.CreateDocumentLinkEventArgs(Me.WorkItem.Document.Name,
        UrlFactory.ToString, Me.WorkItem.Document.LatestVersion.PrimaryContent.MIMEType) With {
        .SourceDocument = Me.WorkItem.Document
        }

      lobjLinkManager = CType(PrimaryConnection.Provider.GetInterface(ProviderClass.LinkManager), ILinkManager)

      Dim lstrLinkId = lobjLinkManager.CreateDocumentLink(lobjCreateDocumentLinkArgs)

      If Not String.IsNullOrEmpty(lstrLinkId) Then
        ProcessedMessage = String.Format("Created Link Object with Id:{0}, LinkUrl:{1}", lstrLinkId, UrlFactory.ToString)
      End If

      If String.IsNullOrEmpty(Me.WorkItem.DestinationDocId) Then
        Me.WorkItem.DestinationDocId = lstrLinkId
      End If

      menuResult = OperationEnumerations.Result.Success

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, ex))
    End Try

    Return menuResult

  End Function

#End Region

End Class
