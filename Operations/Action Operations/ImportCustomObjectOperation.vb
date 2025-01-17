'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ImportCustomObjectOperation.vb
'   Description :  [type_description_here]
'   Created     :  9/2/2015 1:04:07 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Public Class ImportCustomObjectOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "ImportCustomObject"
  'Friend Const PARAM_DELETE_PROPERTIES_WITHOUT_VALUES As String = "DeletePropertiesWithoutValues"
  'Friend Const PARAM_ENFORCE_CLASSIFICATION_COMPLIANCE As String = "EnforceClassificationCompliance"
  'Friend Const PARAM_DOCUMENT_FILING_MODE As String = "DocumentFilingMode"
  'Friend Const PARAM_BASE_PATH_LOCATION As String = "BasePathLocation"
  'Friend Const PARAM_FOLDER_DELIMITER As String = "FolderDelimiter"
  'Friend Const PARAM_LEADING_FOLDER_DELIMITER As String = "LeadingFolderDelimiter"
  'Friend Const PARAM_SET_ANNOTATIONS As String = "SetAnnotations"
  'Friend Const PARAM_SET_PERMISSIONS As String = "SetPermissions"
  'Friend Const PARAM_ADD_AS_MAJOR_VERSION As String = "AddAsMajorVersion"
  'Friend Const PARAM_IMPORT_AS_PACKAGE As String = "ImportAsPackage"
  'Friend Const PARAM_PACKAGE_AS_JSON As String = "PackageAsJson"

  Friend Const PARAM_FOLDER_DELIMITER_DESCRIPTION As String = "Specifies what the folder delimiter is for the destination folder path, the default value is /."

#End Region

#Region "Public Properties"

  'Public ReadOnly Property DeletePropertiesWithoutValues As Boolean
  '  Get
  '    Try
  '      Return GetBooleanParameterValue(PARAM_DELETE_PROPERTIES_WITHOUT_VALUES, True)
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  'End Property

  'Public ReadOnly Property DocumentFilingMode As FilingMode
  '  Get
  '    Try
  '      'Dim lstrFilingMode As String = GetStringParameterValue(PARAM_DOCUMENT_FILING_MODE, FilingMode.UnFiled.ToString)
  '      'Return CType([Enum].Parse(GetType(FilingMode), lstrFilingMode), FilingMode)
  '      Return GetEnumParameterValue(PARAM_DOCUMENT_FILING_MODE, GetType(FilingMode), FilingMode.UnFiled)
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  'End Property

  'Public ReadOnly Property BasePathLocation As ePathLocation
  '  Get
  '    Try
  '      'Dim lstrBasePathLocation As String = GetStringParameterValue(PARAM_BASE_PATH_LOCATION, ePathLocation.Front)
  '      'Return CType([Enum].Parse(GetType(ePathLocation), lstrBasePathLocation), ePathLocation)
  '      Return GetEnumParameterValue(PARAM_BASE_PATH_LOCATION, GetType(ePathLocation), ePathLocation.Front)
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  'End Property

  'Public ReadOnly Property EnforceClassificationCompliance As Boolean
  '  Get
  '    Try
  '      Return GetBooleanParameterValue(PARAM_ENFORCE_CLASSIFICATION_COMPLIANCE, True)
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  'End Property

  'Public ReadOnly Property FolderDelimiter As String
  '  Get
  '    Try
  '      Return GetStringParameterValue(PARAM_FOLDER_DELIMITER, "/")
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  'End Property

  'Public ReadOnly Property LeadingFolderDelimiter As Boolean
  '  Get
  '    Try
  '      Return GetBooleanParameterValue(PARAM_LEADING_FOLDER_DELIMITER, True)
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  'End Property

#End Region

#Region "Public Overrides Methods"

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try

      Return ImportObject()

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

#Region "Protected Methods"

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = New Parameters
      'Dim lstrParameterDescription As String

      'If lobjParameters.Contains(PARAM_SET_PERMISSIONS) = False Then
      '  lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_SET_PERMISSIONS, True, _
      '    "Specifies whether or not to set the new document permissions (if available in the source document).  NOTE: Not all import providers support setting permissions."))
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

  Private Function ImportObject() As Result
    Try

      RunPreOperationChecksForObject(False)

      Dim lobjImporter As ICustomObjectImporter
      Select Case Scope
        Case OperationScope.Source
          lobjImporter = CType(SourceConnection.Provider.GetInterface(ProviderClass.CustomObjectImporter), ICustomObjectImporter)
        Case OperationScope.Destination
          lobjImporter = CType(DestinationConnection.Provider.GetInterface(ProviderClass.CustomObjectImporter), ICustomObjectImporter)
        Case Else
          Throw New InvalidOperationException("Operation scope not set.")
      End Select

      Dim lstrErrorMessage As String = String.Empty

      If String.IsNullOrEmpty(Me.WorkItem.SourceDocId) Then
        Throw New InvalidOperationException("No source object path available")
      End If

      If Me.WorkItem.Object Is Nothing Then
        ' This is an import from a serialized object
        Dim lstrSourcePath As String = Me.WorkItem.SourceDocId.ToLower

        If Not lstrSourcePath.EndsWith(".cof") Then
          Throw New InvalidOperationException(String.Format("Source object path '{0}' does not point to a CTS CustomObject.",
                                                              Me.WorkItem.SourceDocId))
        End If

        If IO.File.Exists(lstrSourcePath) = False Then
          Throw New InvalidOperationException(String.Format("Source object path '{0}' does not point to a CTS CustomObject.",
                                                            Me.WorkItem.SourceDocId))
        End If

        Me.WorkItem.Object = New CustomObject(lstrSourcePath)
      End If

      Dim lobjArgs As New ImportObjectArgs
      lobjArgs.Object = Me.WorkItem.Object

      'lobjArgs.SetAnnotations = GetBooleanParameterValue(PARAM_SET_ANNOTATIONS, True)
      'lobjArgs.SetPermissions = GetBooleanParameterValue(PARAM_SET_PERMISSIONS, True)

      Dim lstrImportedObjectId As String = String.Empty

      If lobjImporter.ImportObject(lobjArgs) = True Then
        menuResult = OperationEnumerations.Result.Success
        Me.WorkItem.DestinationDocId = lobjArgs.Object.Id
      Else
        Me.ProcessedMessage = String.Format("Import Failed: {0}", lobjArgs.ErrorMessage)
        menuResult = OperationEnumerations.Result.Failed
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Me.ProcessedMessage = String.Format("Import Failed: {0}", ex.Message)
      menuResult = OperationEnumerations.Result.Failed
    End Try

    Return menuResult

  End Function

#End Region

End Class
