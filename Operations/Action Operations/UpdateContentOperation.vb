'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  UpdateContentOperation.vb
'   Description :  [type_description_here]
'   Created     :  9/10/2013 9:54:40 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities


#End Region

Public Class UpdateContentOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "UpdateContent"
  ' Friend Const PARAM_CONTENT_ELEMENT_INDEX As String = "ContentElementIndex"
  Friend Const PARAM_NEW_CONTENT_PATH As String = "NewContentPath"
  Private Const DEFAULT_NEW_CONTENT_PATH As String = "C:\Temp\NewContent.pdf"

#End Region

#Region "Public Properties"

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

#End Region

#Region "Friend Methods"

  Friend Overrides Function OnExecute() As Result
    Try
      Return UpdateContent()
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

      'If lobjParameters.Contains(PARAM_CONTENT_ELEMENT_INDEX) = False Then
      '  lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmLong, PARAM_CONTENT_ELEMENT_INDEX, 0, _
      '    "Specifies which content element update.  NOTE: The first content element is specified with a value of zero."))
      'End If

      If lobjParameters.Contains(PARAM_NEW_CONTENT_PATH) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_NEW_CONTENT_PATH,
          DEFAULT_NEW_CONTENT_PATH, "The fully qualified path to a static content file."))
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

  Private Function UpdateContent() As Result
    Try

      RunPreOperationChecks(True)

      Dim lstrNewContentPath As String = GetStringParameterValue(PARAM_NEW_CONTENT_PATH, DEFAULT_NEW_CONTENT_PATH)

      If IO.File.Exists(lstrNewContentPath) = False Then
        Throw New InvalidPathException(lstrNewContentPath)
      End If

      Dim lobjWorkingVersion As Version = Me.WorkItem.Document.LatestVersion

      With lobjWorkingVersion.Contents
        .Clear()
        .Add(lstrNewContentPath)
      End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, ex))
    End Try

    Return menuResult

  End Function

#End Region

End Class
