' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  UpdateOperation.vb
'  Description :  [type_description_here]
'  Created     :  12/14/2011 7:34:00 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Public Class UpdatePropertiesOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "UpdateProperties"
  Friend Const PARAM_VERSION_SCOPE As String = "VersionScope"
  Friend Const PARAM_PROPERTIES_TO_UPDATE As String = "PropertiesToUpdate"

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

      Return UpdateDocument()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Protected Methods"

  Public Overrides Sub CheckParameters()
    Try
      UpdateParameterToEnum(PARAM_VERSION_SCOPE, GetType(VersionScopeEnum))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try

      Dim lobjParameters As IParameters = New Parameters

      If lobjParameters.Contains(PARAM_VERSION_SCOPE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_VERSION_SCOPE, VersionScopeEnum.MostCurrentVersion,
          GetType(VersionScopeEnum),
          "Specifies which versions of the document should be updated (NOTE: Selective versions are not supported on all providers."))
      End If

      If lobjParameters.Contains(PARAM_PROPERTIES_TO_UPDATE) = False Then
        'Dim lobjFolderPathsParam As New Parameter
        Dim lobjFolderPathsParam As IParameter = ParameterFactory.Create(PropertyType.ecmString,
          PARAM_PROPERTIES_TO_UPDATE, Cardinality.ecmMultiValued)
        With lobjFolderPathsParam
          .SystemName = PARAM_PROPERTIES_TO_UPDATE
          .Description = "The list of properties to update."
          .Values.Add("DocumentTitle")
        End With
        lobjParameters.Add(lobjFolderPathsParam)
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

  Private Function UpdateDocument() As Result

    Dim lobjBCSProvider As IBasicContentServicesProvider

    Try

      RunPreOperationChecks(True)

      lobjBCSProvider = CType(PrimaryConnection.Provider.GetInterface(ProviderClass.BasicContentServices),
        IBasicContentServicesProvider)

      Dim lobjPropertyNamesParam As IParameter = CType(Parameters, Parameters)(PARAM_PROPERTIES_TO_UPDATE)
      Dim lobjPropertyNames As New List(Of String)
      If lobjPropertyNamesParam.HasValue Then
        For Each lobjFolderPath As Object In CType(lobjPropertyNamesParam.Values, IEnumerable(Of Object))
          lobjPropertyNames.Add(lobjFolderPath.ToString)
        Next
      End If

      If (lobjBCSProvider IsNot Nothing) Then
        Dim lobjVersionScope As VersionScopeEnum = GetEnumParameterValue(PARAM_VERSION_SCOPE, GetType(VersionScopeEnum),
          VersionScopeEnum.MostCurrentVersion)
        Dim lobjDocumentPropertyArgs As DocumentPropertyArgs = Nothing

        Dim lobjMasterProperties As New ECMProperties
        Dim lobjSelectedProperties As New ECMProperties
        Dim lobjSelectedProperty As ECMProperty = Nothing


        Select Case lobjVersionScope
          Case VersionScopeEnum.MostCurrentVersion
            lobjSelectedProperties = GetSelectedProperties(Me.WorkItem.Document.LatestVersion.Properties, lobjPropertyNames)
            lobjDocumentPropertyArgs = New DocumentPropertyArgs(Me.WorkItem.Document.ID, lobjVersionScope, lobjSelectedProperties)
            menuResult = ConvertResult(lobjBCSProvider.UpdateDocumentProperties(lobjDocumentPropertyArgs))

          Case VersionScopeEnum.FirstVersion

            lobjSelectedProperties = GetSelectedProperties(Me.WorkItem.Document.FirstVersion.Properties, lobjPropertyNames)
            lobjDocumentPropertyArgs = New DocumentPropertyArgs(Me.WorkItem.Document.ID, lobjVersionScope, lobjSelectedProperties)
            menuResult = ConvertResult(lobjBCSProvider.UpdateDocumentProperties(lobjDocumentPropertyArgs))

            'Case VersionScopeEnum.AllVersions
            '  'For Each lobjVersion As Version In Me.WorkItem.Document.Versions
            '  '  lobjDocumentPropertyArgs = New DocumentPropertyArgs(Me.WorkItem.Document.ID, lobjVersion.ID.ToString, _
            '  '                                 lobjVersion.Properties)
            '  '  menuResult = ConvertResult(lobjBCSProvider.UpdateDocumentProperties(lobjDocumentPropertyArgs))
            '  '  If menuResult = OperationEnumerations.Result.Failed Then
            '  '    Exit For
            '  '  End If
            '  'Next
            '  For Each lobjVersion As Version In Me.WorkItem.Document.Versions
            '    lobjSelectedProperties = GetSelectedProperties(lobjVersion.Properties, lobjPropertyNames)
            '    lobjDocumentPropertyArgs = New DocumentPropertyArgs(Me.WorkItem.Document.ID, lobjVersion.ID.ToString, lobjSelectedProperties)
            '    menuResult = ConvertResult(lobjBCSProvider.UpdateDocumentProperties(lobjDocumentPropertyArgs))
            '    If menuResult = OperationEnumerations.Result.Failed Then
            '      Exit For
            '    End If
            '  Next
          Case Else
            Throw New InvalidVersionSpecificationException(Me.WorkItem.Document,
              String.Format("Version scope {0} not supported for update properties operation.", lobjVersionScope.ToString))
        End Select

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

  Private Shared Function GetSelectedProperties(lpMasterProperties As ECMProperties, lpTargetPropertyNames As List(Of String)) As ECMProperties
    Try
      Dim lobjSelectedProperties As New ECMProperties
      Dim lobjSelectedProperty As ECMProperty = Nothing

      If lpTargetPropertyNames.Count > 0 Then
        For Each lstrPropertyName As String In lpTargetPropertyNames
          If lpMasterProperties.PropertyExists(lstrPropertyName, False, lobjSelectedProperty) Then
            lobjSelectedProperties.Add(lobjSelectedProperty)
          End If
        Next
      Else
        lobjSelectedProperties = lpMasterProperties
      End If

      Return lobjSelectedProperties

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
