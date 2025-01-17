' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  UpdatePropertyOperation.vb
'  Description :  [type_description_here]
'  Created     :  11/17/2015 9:07:23 AM
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

Public Class UpdatePropertyOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "UpdateProperty"
  Friend Const PARAM_VERSION_SCOPE As String = "VersionScope"
  Friend Const PARAM_PROPERTY_TO_UPDATE As String = "PropertyToUpdate"
  Friend Const PARAM_NEW_PROPERTY_VALUE As String = "NewPropertyValue"

#End Region

#Region "Public Properties"

  Public Overrides ReadOnly Property CanRollback As Boolean
    Get
      Return False
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

      If lobjParameters.Contains(PARAM_PROPERTY_TO_UPDATE) = False Then
        'Dim lobjFolderPathsParam As New Parameter
        Dim lobjTargetPropertyParam As IParameter = ParameterFactory.Create(PropertyType.ecmString,
          PARAM_PROPERTY_TO_UPDATE, Cardinality.ecmSingleValued)
        With lobjTargetPropertyParam
          .SystemName = PARAM_PROPERTY_TO_UPDATE
          .Description = "The property to update."
          If Me.PrimaryConnection IsNot Nothing Then
            Try
              Dim lobjSourceRepository As Repository = Nothing
              If Me.PrimaryConnection.Repository Is Nothing Then
                lobjSourceRepository = Me.PrimaryConnection.Repository
              Else
                lobjSourceRepository = Me.PrimaryConnection.GetRepository
              End If

              If lobjSourceRepository IsNot Nothing Then
                Dim lobjStandardValues As New List(Of String)
                For Each lobjProperty As ClassificationProperty In lobjSourceRepository.Properties
                  If lobjProperty.Settability = ClassificationProperty.SettabilityEnum.READ_WRITE Then
                    lobjStandardValues.Add(lobjProperty.Name)
                  End If
                Next
              End If
            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              ' Just forget it and move on.
            End Try
          End If
        End With
        lobjParameters.Add(lobjTargetPropertyParam)
      End If

      If lobjParameters.Contains(PARAM_NEW_PROPERTY_VALUE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_NEW_PROPERTY_VALUE,
          String.Empty, "Specifies the new value for the property."))
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

      RunPreOperationChecks(False)

      Dim lstrDocumentId As String = Nothing

      If Me.Scope = OperationScope.Source Then
        lstrDocumentId = Me.WorkItem.SourceDocId
      Else
        lstrDocumentId = Me.WorkItem.DestinationDocId
      End If

      lobjBCSProvider = CType(PrimaryConnection.Provider.GetInterface(ProviderClass.BasicContentServices),
        IBasicContentServicesProvider)

      Dim lobjPropertyNameParam As IParameter = CType(Parameters, Parameters)(PARAM_PROPERTY_TO_UPDATE)
      Dim lobjNewPropertyValueParam As IParameter = CType(Parameters, Parameters)(PARAM_NEW_PROPERTY_VALUE)

      Dim lstrPropertyName As String = Nothing
      Dim lstrNewPropertyValue As String

      If lobjPropertyNameParam.HasValue Then
        lstrPropertyName = lobjPropertyNameParam.Value
      End If

      If lobjNewPropertyValueParam.HasValue Then
        lstrNewPropertyValue = lobjNewPropertyValueParam.Value
      Else
        lstrNewPropertyValue = String.Empty
      End If

      If (lobjBCSProvider IsNot Nothing) Then
        Dim lobjVersionScope As VersionScopeEnum = GetEnumParameterValue(PARAM_VERSION_SCOPE, GetType(VersionScopeEnum),
          VersionScopeEnum.MostCurrentVersion)
        Dim lobjDocumentPropertyArgs As DocumentPropertyArgs = Nothing

        Dim lobjMasterProperties As New ECMProperties
        Dim lobjSelectedProperties As New ECMProperties
        Dim lobjSelectedProperty As ECMProperty = Nothing


        'Select Case lobjVersionScope
        '  Case VersionScopeEnum.MostCurrentVersion
        '    lobjSelectedProperty = GetSelectedProperty(Me.WorkItem.Document.LatestVersion.Properties, lstrPropertyName)

        '  Case VersionScopeEnum.FirstVersion

        '    lobjSelectedProperty = GetSelectedProperty(Me.WorkItem.Document.FirstVersion.Properties, lstrPropertyName)

        '  Case Else
        '    Throw New Exceptions.InvalidVersionSpecificationException(Me.WorkItem.Document, _
        '      String.Format("Version scope {0} not supported for update properties operation.", lobjVersionScope.ToString))
        'End Select

        Select Case lobjVersionScope
          Case VersionScopeEnum.MostCurrentVersion, VersionScopeEnum.CurrentReleasedVersion, VersionScopeEnum.AllVersions
            ' We are good

          Case Else
            Throw New InvalidVersionSpecificationException(lstrDocumentId,
              String.Format("Version scope {0} not supported for update property operation.", lobjVersionScope.ToString))
        End Select

        lobjSelectedProperty = GetSelectedProperty(lstrPropertyName)
        lobjSelectedProperty.ChangePropertyValue(lstrNewPropertyValue)

        lobjSelectedProperties.Add(lobjSelectedProperty)
        lobjDocumentPropertyArgs = New DocumentPropertyArgs(lstrDocumentId, lobjVersionScope, lobjSelectedProperties)
        menuResult = ConvertResult(lobjBCSProvider.UpdateDocumentProperties(lobjDocumentPropertyArgs))

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

  Private Function GetSelectedProperty(lpTargetPropertyName As String) As ECMProperty
    Try

      Dim lobjSourceRepository As Repository

      If Me.SourceConnection.Repository IsNot Nothing Then
        lobjSourceRepository = Me.SourceConnection.Repository
      Else
        lobjSourceRepository = Me.SourceConnection.GetRepository
      End If

      Dim lobjSourceClassificationProperties As ClassificationProperties = lobjSourceRepository.Properties
      Dim lobjSelectedProperty As ClassificationProperty = Nothing
      Dim lobjTargetProperty As ECMProperty = Nothing

      If Not lobjSourceClassificationProperties.PropertyExists(lpTargetPropertyName, False) Then
        Throw New PropertyDoesNotExistException(lpTargetPropertyName)
      Else
        lobjSelectedProperty = lobjSourceClassificationProperties.Item(lpTargetPropertyName)
        lobjTargetProperty = PropertyFactory.Create(lobjSelectedProperty)
        Return lobjTargetProperty
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetSelectedProperty(lpMasterProperties As ECMProperties, lpTargetPropertyName As String) As ECMProperty
    Try
      Dim lobjSelectedProperty As ECMProperty = Nothing

      If Not lpMasterProperties.PropertyExists(lpTargetPropertyName, False, lobjSelectedProperty) Then
        Throw New PropertyDoesNotExistException(lpTargetPropertyName)
      End If

      Return lobjSelectedProperty

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
