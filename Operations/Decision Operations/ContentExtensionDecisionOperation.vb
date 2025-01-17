' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ContentExtensionDecisionOperation.vb
'  Description :  [type_description_here]
'  Created     :  4/13/2012 11:28:21 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class ContentExtensionDecisionOperation
  Inherits DecisionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "ContentExtensionDecision"
  Friend Const PARAM_EXTENSIONS As String = "Extensions"

  Public Const MODE_VALID As String = "Valid"
  Public Const MODE_INVALID As String = "Invalid"

#End Region

#Region "Clas Variables"

  Private WithEvents MobjExtensions As ObservableCollection(Of String) = Nothing

#End Region

#Region "Public Properties"

  Protected ReadOnly Property Extensions As ObservableCollection(Of String)
    Get
      Try
        If MobjExtensions Is Nothing Then
          MobjExtensions = GetExtensions()
        End If
        Return MobjExtensions
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

#End Region

#Region "IDecisionOperation Implementation"

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

  Public Overrides Function Evaluate() As Boolean
    Try
      Return EvaluateExtension()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  'Public Function Clone() As Object Implements System.ICloneable.Clone
  '  Try
  '    Throw New NotImplementedException
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

#End Region

#Region "Protected Methods"

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = MyBase.GetDefaultParameters

      Dim lobjValidExtensionsParameter As IParameter = ParameterFactory.Create(PropertyType.ecmString, PARAM_EXTENSIONS, Cardinality.ecmMultiValued)
      lobjValidExtensionsParameter.Description = "The list of extensions to test for."
      lobjParameters.Add(lobjValidExtensionsParameter)

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Function GetExtensions() As ObservableCollection(Of String)
    Try
      Dim lobjExtensions As New ObservableCollection(Of String)

      Dim lobjExtensionValues As Values = CType(Me.Parameters.Item(PARAM_EXTENSIONS).Values, Values)
      Dim lstrSmallExtension As String = String.Empty

      For Each lstrExtension As String In lobjExtensionValues
        lstrSmallExtension = lstrExtension.ToLower
        If lobjExtensions.Contains(lstrSmallExtension) = False Then
          lobjExtensions.Add(lstrSmallExtension)
        End If
      Next

      Return lobjExtensions

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function EvaluateExtension() As Boolean
    Try

      Dim lobjExtensions As ObservableCollection(Of String) = GetExtensions()
      Dim lenuVersionScope As VersionScopeEnum = CType([Enum].Parse(lenuVersionScope.GetType, CStr(Me.Parameters.Item(PARAM_VERSION_SCOPE).Value)), VersionScopeEnum)
      Dim lintContentElementIndex As Integer = CInt(Me.Parameters.Item(PARAM_CONTENT_ELEMENT_INDEX).Value)
      Dim lenuMode As ModeEnum = CType([Enum].Parse(lenuMode.GetType, CStr(Me.Parameters.Item(PARAM_MODE).Value)), ModeEnum)
      Dim lblnVersionResult As Boolean

      Select Case lenuVersionScope
        Case VersionScopeEnum.FirstVersion
          Return EvaluateVersionExtension(Me.WorkItem.Document.FirstVersion, lintContentElementIndex, lenuMode)
        Case VersionScopeEnum.MostCurrentVersion, VersionScopeEnum.CurrentReleasedVersion
          Return EvaluateVersionExtension(Me.WorkItem.Document.LatestVersion, lintContentElementIndex, lenuMode)
        Case VersionScopeEnum.AllVersions
          For Each lobjVersion As Version In Me.WorkItem.Document.Versions
            lblnVersionResult = EvaluateVersionExtension(lobjVersion, lintContentElementIndex, lenuMode)
            If lblnVersionResult = False Then
              Return lblnVersionResult
            End If
          Next
          Return lblnVersionResult
        Case VersionScopeEnum.FirstNVersions, VersionScopeEnum.LastNVersions
          Throw New ArgumentOutOfRangeException(lenuVersionScope.ToString, "Version ranges outside of all versions are not supported for this operation.")
        Case Else
          Throw New ArgumentOutOfRangeException(lenuVersionScope.ToString, "Invalid version scope.")
      End Select

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      'LogSession.LogException("An exception occured in 'EvaluateExtension'.", ex)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Private Function EvaluateVersionExtension(ByVal lpVersion As Version, ByVal lpContentIndex As Integer, ByVal lpMode As ModeEnum) As Boolean
    Try
      If lpContentIndex > -1 Then
        Return EvaluateContentExtension(lpVersion.Contents(lpContentIndex), lpMode)
      Else
        Throw New ArgumentOutOfRangeException(NameOf(lpContentIndex), "Only positive content index values for specific content elements are supported.")
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      'LogSession.LogException("An exception occured in 'EvaluateVersionExtension'.", ex)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function EvaluateContentExtension(ByVal lpContent As Content, ByVal lpMode As ModeEnum) As Boolean
    Try
      Dim lstrExtension As String = lpContent.FileExtension.ToLower

      Select Case lpMode
        Case ModeEnum.Valid
          If Extensions.Contains(lstrExtension) Then
            Return True
          Else
            Return False
          End If
        Case ModeEnum.Invalid
          If Extensions.Contains(lstrExtension) Then
            Return False
          Else
            Return True
          End If
        Case Else ' We should never get here
          If Extensions.Contains(lstrExtension) Then
            Return False
          Else
            Return True
          End If
      End Select
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      'LogSession.LogException("An exception occured in 'EvaluateContentExtension'.", ex)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

  Private Sub MobjExtensions_CollectionChanged(sender As Object, e As Specialized.NotifyCollectionChangedEventArgs) Handles MobjExtensions.CollectionChanged
    Try
      Dim lobjExtensionValues As Values = CType(Me.Parameters.Item(PARAM_EXTENSIONS).Values, Values)

      Select Case e.Action
        Case NotifyCollectionChangedAction.Add
          For Each lobjItem As Object In e.NewItems
            lobjExtensionValues.Add(lobjItem)
          Next

        Case NotifyCollectionChangedAction.Remove
          lobjExtensionValues.Remove(e.OldItems)

        Case NotifyCollectionChangedAction.Replace
          For lintItemCounter As Integer = 0 To e.NewItems.Count - 1
            lobjExtensionValues.Remove(lobjExtensionValues.GetItemByName(e.OldItems(lintItemCounter)))
            lobjExtensionValues.Add(e.NewItems(lintItemCounter))
          Next

        Case NotifyCollectionChangedAction.Reset

      End Select
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

End Class
