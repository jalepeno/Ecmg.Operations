' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ContentSizeDecisionOperation.vb
'  Description :  [type_description_here]
'  Created     :  4/19/2012 9:04:47 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class ContentSizeDecisionOperation
  Inherits DecisionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "ContentSizeDecision"

  Friend Const PARAM_MINIMUM_CONTENT_SIZE As String = "MinimumContentSizeKB"
  Friend Const PARAM_MAXIMUM_CONTENT_SIZE As String = "MaximumContentSizeKB"

#End Region

#Region "IDecisionOperation Implementation"

  Public Overrides Function Evaluate() As Boolean
    Try
      Return EvaluateContentSize()
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

#End Region

#Region "Protected Methods"

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = MyBase.GetDefaultParameters

      If lobjParameters.Contains(PARAM_MINIMUM_CONTENT_SIZE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmLong, PARAM_MINIMUM_CONTENT_SIZE, 0,
          "The minimum content size in kilobytes."))
      End If

      If lobjParameters.Contains(PARAM_MAXIMUM_CONTENT_SIZE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmLong, PARAM_MAXIMUM_CONTENT_SIZE, 102400,
          "The maximum content size in kilobytes."))
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

  Private Function EvaluateContentSize() As Boolean
    Try

      Dim lenuVersionScope As VersionScopeEnum = CType([Enum].Parse(lenuVersionScope.GetType,
                                                                    CStr(Me.Parameters.Item(PARAM_VERSION_SCOPE).Value)), VersionScopeEnum)
      Dim lintContentElementIndex As Integer = CInt(Me.Parameters.Item(PARAM_CONTENT_ELEMENT_INDEX).Value)
      Dim lblnVersionResult As Boolean

      Select Case lenuVersionScope
        Case VersionScopeEnum.FirstVersion
          lblnVersionResult = EvaluateVersionSize(Me.WorkItem.Document.FirstVersion, lintContentElementIndex)
        Case VersionScopeEnum.MostCurrentVersion, VersionScopeEnum.CurrentReleasedVersion
          lblnVersionResult = EvaluateVersionSize(Me.WorkItem.Document.LatestVersion, lintContentElementIndex)
        Case VersionScopeEnum.AllVersions
          For Each lobjVersion As Version In Me.WorkItem.Document.Versions
            lblnVersionResult = EvaluateVersionSize(lobjVersion, lintContentElementIndex)
            If lblnVersionResult = False Then
              Return lblnVersionResult
            End If
          Next
      End Select

      Return lblnVersionResult

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function EvaluateVersionSize(ByVal lpVersion As Version, ByVal lpContentIndex As Integer) As Boolean
    Try
      If lpContentIndex > -1 Then

        Dim lintMinFileSize As Integer = CInt(Me.Parameters.Item(PARAM_MINIMUM_CONTENT_SIZE).Value)
        Dim lintMaxFileSize As Integer = CInt(Me.Parameters.Item(PARAM_MAXIMUM_CONTENT_SIZE).Value)
        Dim ldblContentFileSize As Double = lpVersion.Contents(lpContentIndex).FileSize.Kilobytes

        If (ldblContentFileSize < lintMinFileSize) OrElse (ldblContentFileSize > lintMaxFileSize) Then
          Return False
        Else
          Return True
        End If
      Else
        Throw New ArgumentOutOfRangeException(NameOf(lpContentIndex), "Only positive content index values for specific content elements are supported.")
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
