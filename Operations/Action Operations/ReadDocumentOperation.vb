' ********************************************************************************
' '  Document    :  ReadDocumentOperation.vb
' '  Description :  [type_description_here]
' '  Created     :  10/10/2012-13:38:17
' '  <copyright company="ECMG">
' '      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class ReadDocumentOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "ReadDocument"
  Private Const PARAM_SCOPE As String = "Scope"

#End Region

#Region "Public Overrides Methods"

  Public Overrides ReadOnly Property Name As String
    Get
      Try
        Return OPERATION_NAME
      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
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

  Friend Overrides Function OnExecute() As Result
    Try
      Dim lstrScope As String = GetStringParameterValue(PARAM_SCOPE, "Source")
      Scope = CType([Enum].Parse(GetType(OperationScope), lstrScope), OperationScope)

      Dim lstrSourcePath As String = Nothing
      Select Case Scope
        Case OperationScope.Source
          lstrSourcePath = Me.WorkItem.SourceDocId.ToLower
        Case OperationScope.Destination
          lstrSourcePath = Me.WorkItem.DestinationDocId.ToLower
      End Select

      If Not lstrSourcePath.EndsWith(".cpf") Then
        If Not lstrSourcePath.EndsWith(".cdf") Then
          Throw New InvalidOperationException(String.Format("Document path '{0}' does not point to a CTS Document.",
                                                            Me.WorkItem.SourceDocId))
        End If
      End If

      If IO.File.Exists(lstrSourcePath) = False Then
        Throw New InvalidOperationException(String.Format("Document path '{0}' does not point to a CTS Document.",
                                                          Me.WorkItem.SourceDocId))
      End If

      Me.WorkItem.Document = New Document(lstrSourcePath)
      menuResult = OperationEnumerations.Result.Success
      Return menuResult

    Catch Ex As Exception
      ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
      Me.ProcessedMessage = String.Format("Read Document Failed: {0}", Ex.Message)
      menuResult = OperationEnumerations.Result.Failed
      Return menuResult
    End Try
  End Function

#End Region

#Region "Protected Methods"

  Public Overrides Sub CheckParameters()
    Try
      UpdateParameterToEnum(PARAM_SCOPE, GetType(OperationScope))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = New Parameters

      If lobjParameters.Contains(PARAM_SCOPE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_SCOPE, OperationScope.Source,
          GetType(OperationScope), "Specifies whether the source or destination document will be read."))
      End If

      Return lobjParameters

    Catch Ex As Exception
      ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class