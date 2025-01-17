'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IsPdfDecisionOperation.vb
'   Description :  [type_description_here]
'   Created     :  9/9/2013 8:34:00 AM
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

Public Class IsPdfDecisionOperation
  Inherits DecisionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "IsPdfDecision"

#End Region

#Region "IDecisionOperation Implementation"

  Public Overrides Function Evaluate() As Boolean
    Try
      Return EvaluateForPdf()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function EvaluateForPdf() As Boolean
    Try
      Dim lenuVersionScope As VersionScopeEnum = CType([Enum].Parse(lenuVersionScope.GetType, CStr(Me.Parameters.Item(PARAM_VERSION_SCOPE).Value)), VersionScopeEnum)
      Dim lenuMode As ModeEnum = CType([Enum].Parse(lenuMode.GetType, CStr(Me.Parameters.Item(PARAM_MODE).Value)), ModeEnum)
      Dim lblnVersionResult As Boolean

      Select Case lenuVersionScope
        Case VersionScopeEnum.FirstVersion
          Return EvaluateVersionForPdf(Me.WorkItem.Document.FirstVersion, lenuMode)
        Case VersionScopeEnum.MostCurrentVersion, VersionScopeEnum.CurrentReleasedVersion
          Return EvaluateVersionForPdf(Me.WorkItem.Document.LatestVersion, lenuMode)
        Case VersionScopeEnum.AllVersions
          For Each lobjVersion As Version In Me.WorkItem.Document.Versions
            lblnVersionResult = EvaluateVersionForPdf(lobjVersion, lenuMode)
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

  Private Function EvaluateVersionForPdf(ByVal lpVersion As Version, ByVal lpMode As ModeEnum) As Boolean
    Try
      Dim lblnReturnValue As Boolean

      If lpVersion.Contents Is Nothing OrElse lpVersion.Contents.Count = 0 Then
        lblnReturnValue = False
      Else
        lblnReturnValue = Helper.IsPdfStream(lpVersion.PrimaryContent.ToMemoryStream)
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

End Class
