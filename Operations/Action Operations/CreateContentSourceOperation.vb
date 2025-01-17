'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CreateContentSource.vb
'   Description :  [type_description_here]
'   Created     :  2/20/2013 10:10:00 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.Specialized
Imports System.Text
Imports Documents.Configuration
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Public Class CreateContentSourceOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "CreateContentSource"
  Friend Const PARAM_CONTENT_SOURCE_NAME As String = "ContentSourceName"
  Friend Const PARAM_PROVIDER_NAME As String = "ProviderName"
  Friend Const PARAM_REPLACE_EXISTING As String = "ReplaceExisting"
  Friend Const PARAM_PROPERTY_LIST As String = "PropertyList"
  Friend Const DEFAULT_CONTENT_SOURCE_NAME As String = "ContentSource Name"
  Friend Const DEFAULT_PROVIDER_NAME As String = "Provider Name"

  Private Const CONTENT_SOURCE_NAME_PLACEHOLDER As String = "{ContentSourceName}"

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
      Return CreateContentSource()
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
          "The name of the new content source."))
      End If

      If lobjParameters.Contains(PARAM_PROVIDER_NAME) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_PROVIDER_NAME, DEFAULT_PROVIDER_NAME,
          "The provider to be used for the new content source."))
      End If

      If lobjParameters.Contains(PARAM_REPLACE_EXISTING) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_REPLACE_EXISTING, False,
          "Specifies whether or not to replace any existing content source with the same name."))
      End If

      If lobjParameters.Contains(PARAM_PROPERTY_LIST) = False Then
        lobjParameters.Add(New ObjectParameter(PropertyType.ecmObject, PARAM_PROPERTY_LIST,
          CreateSamplePropertyList(), "The dictionary of properties used to define the new content source."))
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

  Private Shared Function CreateSamplePropertyList() As StringDictionary
    Try
      Dim lobjDictionary As New StringDictionary
      Dim lobjStandardBaseProperties As ProviderProperties = CProvider.GetAllStandardProviderProperties

      For Each lobjProperty As ProviderProperty In lobjStandardBaseProperties
        If lobjProperty.Name.Equals("exportpath", StringComparison.CurrentCultureIgnoreCase) Then
          lobjProperty.Value = String.Format("%CtsDocsPath%\Exports\{0}", CONTENT_SOURCE_NAME_PLACEHOLDER)
        ElseIf lobjProperty.Name.Equals("importpath", StringComparison.CurrentCultureIgnoreCase) Then
          lobjProperty.Value = String.Format("%CtsDocsPath%\Imports\{0}", CONTENT_SOURCE_NAME_PLACEHOLDER)
        End If
        If lobjProperty.HasValue Then
          lobjDictionary.Add(lobjProperty.Name, lobjProperty.Value.ToString)
        ElseIf lobjProperty.DefaultValue IsNot Nothing Then
          lobjDictionary.Add(lobjProperty.Name, lobjProperty.DefaultValue.ToString)
        Else
          lobjDictionary.Add(lobjProperty.Name, String.Empty)
        End If
      Next

      Return lobjDictionary

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateContentSource() As OperationEnumerations.Result
    Try

      mstrContentSourceName = GetStringParameterValue(PARAM_CONTENT_SOURCE_NAME, DEFAULT_CONTENT_SOURCE_NAME)

      If ConnectionSettings.Instance.ContentSourceNames.Contains(mstrContentSourceName) Then
        Dim lblnReplaceExisting As Boolean = GetBooleanParameterValue(PARAM_REPLACE_EXISTING, False)
        If lblnReplaceExisting = False Then
          Return OperationEnumerations.Result.PreviouslySucceeded
        End If
      End If

      Dim lstrContentSourceConnectionString As String = CreateConnectionString()

      If menuResult = OperationEnumerations.Result.Failed Then
        Return menuResult
      End If

      If String.IsNullOrEmpty(lstrContentSourceConnectionString) Then
        menuResult = OperationEnumerations.Result.Failed
        Return menuResult
      End If

      Dim lobjContentSource As New ContentSource(lstrContentSourceConnectionString)

      ConnectionSettings.Instance.ContentSourceConnectionStrings.Add(lstrContentSourceConnectionString)
      ConnectionSettings.Instance.Save()

      menuResult = OperationEnumerations.Result.Success
      Me.ProcessedMessage = String.Format("Created content source '{0}'", mstrContentSourceName)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, ex))
    End Try

    Return menuResult

  End Function

  Private Function CreateConnectionString() As String
    Try

      Dim lstrProviderName As String = GetStringParameterValue(PARAM_PROVIDER_NAME, DEFAULT_PROVIDER_NAME)
      Dim lobjPropertyList As StringDictionary = CType(GetParameterValue(PARAM_PROPERTY_LIST, New StringDictionary), StringDictionary)
      Dim lobjStringBuilder As New StringBuilder

      If String.IsNullOrEmpty(mstrContentSourceName) Then
        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to create content source without a valid content source name."))
        Return String.Empty
      End If

      If String.Equals(mstrContentSourceName, DEFAULT_CONTENT_SOURCE_NAME) Then
        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to create content source with the default content source name."))
        Return String.Empty
      End If

      If String.IsNullOrEmpty(lstrProviderName) Then
        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to create content source without a valid provider name."))
        Return String.Empty
      End If

      If String.Equals(lstrProviderName, DEFAULT_CONTENT_SOURCE_NAME) Then
        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, "Unable to create content source with the default provider name."))
        Return String.Empty
      End If

      lobjStringBuilder.AppendFormat("Name={0};", mstrContentSourceName)
      lobjStringBuilder.AppendFormat("Provider={0};", lstrProviderName)

      For Each lstrKey As String In lobjPropertyList.Keys
        Select Case lstrKey.ToLower
          Case "exportpath", "importpath"
            lobjStringBuilder.AppendFormat("{0}={1};", lstrKey, lobjPropertyList(lstrKey).Replace(CONTENT_SOURCE_NAME_PLACEHOLDER, mstrContentSourceName))
          Case Else
            lobjStringBuilder.AppendFormat("{0}={1};", lstrKey, lobjPropertyList(lstrKey))
        End Select
      Next

      lobjStringBuilder.Remove(lobjStringBuilder.Length - 1, 1)

      Return lobjStringBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
