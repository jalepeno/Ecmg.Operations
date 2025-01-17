'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  MultiContentDecisionOperation.vb
'   Description :  [type_description_here]
'   Created     :  9/9/2013 8:07:52 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class MultiContentDecisionOperation
  Inherits DecisionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "MultiContentDecision"

#End Region

#Region "Protected Methods"

  'Protected Overrides Function GetDefaultParameters() As IParameters
  '  Try
  '    Dim lobjParameters As IParameters = New Parameters

  '    If lobjParameters.Contains(PARAM_VERSION_SCOPE) = False Then
  '      lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_VERSION_SCOPE, _
  '        VersionScopeEnum.AllVersions, GetType(VersionScopeEnum), _
  '        "Specifies which version(s) to use for the evaluation."))
  '    End If

  '    If lobjParameters.Contains(PARAM_CONTENT_ELEMENT_INDEX) = False Then
  '      lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmLong, PARAM_CONTENT_ELEMENT_INDEX, 0, _
  '        "Specifies which content element to use for the evaluation.  NOTE: The first content element is specified with a value of zero."))
  '    End If

  '    If lobjParameters.Contains(PARAM_MODE) = False Then
  '      lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_MODE, ModeEnum.Valid, GetType(ModeEnum), _
  '        "Specifies whether the mode is valid or invalid.  When the mode is valid, the evaluation will return true if the extension exists, when the mode is invalid, the evaluation will return false if the extension exists."))
  '    End If

  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

#End Region

#Region "IDecisionOperation Implementation"

  Public Overrides Function Evaluate() As Boolean
    Try
      Return EvaluateForMultiContent()
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

  Private Function EvaluateForMultiContent() As Boolean
    Try
      Dim lenuVersionScope As VersionScopeEnum = CType([Enum].Parse(lenuVersionScope.GetType, CStr(Me.Parameters.Item(PARAM_VERSION_SCOPE).Value)), VersionScopeEnum)
      Dim lenuMode As ModeEnum = CType([Enum].Parse(lenuMode.GetType, CStr(Me.Parameters.Item(PARAM_MODE).Value)), ModeEnum)
      Dim lblnVersionResult As Boolean

      Select Case lenuVersionScope
        Case VersionScopeEnum.FirstVersion
          Return EvaluateVersionForMultiContent(Me.WorkItem.Document.FirstVersion, lenuMode)
        Case VersionScopeEnum.MostCurrentVersion, VersionScopeEnum.CurrentReleasedVersion
          Return EvaluateVersionForMultiContent(Me.WorkItem.Document.LatestVersion, lenuMode)
        Case VersionScopeEnum.AllVersions
          For Each lobjVersion As Version In Me.WorkItem.Document.Versions
            lblnVersionResult = EvaluateVersionForMultiContent(lobjVersion, lenuMode)
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
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Shared Function EvaluateVersionForMultiContent(ByVal lpVersion As Version, ByVal lpMode As ModeEnum) As Boolean
    Try

      Dim lblnReturnValue As Boolean

      If lpVersion.Contents Is Nothing OrElse lpVersion.Contents.Count < 2 Then
        lblnReturnValue = False
      Else
        lblnReturnValue = True
      End If

      If lpMode = ModeEnum.Invalid Then
        lblnReturnValue = Not lblnReturnValue
      End If

      Return lblnReturnValue

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
