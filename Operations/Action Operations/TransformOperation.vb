' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  TransformationOperation.vb
'  Description :  [type_description_here]
'  Created     :  11/30/2011 1:22:38 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Transformations
Imports Documents.Utilities

#End Region

Public Class TransformOperation
  Inherits ActionOperation

#Region "Class Constants"

  Public Const OPERATION_NAME As String = "Transform"
  Friend Const PARAM_ROOT_TRANSFORMATION As String = "RootTransformation"

#End Region

#Region "Class Variables"

  Private WithEvents mobjTransformations As New TransformationCollection
  Private mstrRootTransformation As String = String.Empty

#End Region

#Region "Public Properties"

  Public ReadOnly Property Transformations As TransformationCollection
    Get
      Return mobjTransformations
    End Get
  End Property

  Public ReadOnly Property RootTransformation As Transformation
    Get
      Try
        Return Transformations.PrimaryTransformation
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property RootTransformationName As String
    Get
      Return mstrRootTransformation
    End Get
  End Property

#End Region

#Region "Public Overrides Methods"

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try

      Return TransformDocument()

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
      Return False
    End Get
  End Property

#End Region

#Region "Protected Methods"

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try
      Dim lobjParameters As IParameters = New Parameters

      If lobjParameters.Contains(PARAM_ROOT_TRANSFORMATION) = False Then
        ' lobjParameters.Add(New Parameter(PropertyType.ecmString, PARAM_ROOT_TRANSFORMATION, String.Empty))
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_ROOT_TRANSFORMATION, String.Empty,
          "The name of the associated transformation to use."))
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

  Private Function TransformDocument() As Result
    Try

      RunPreOperationChecks(True)
      ' ApplicationLogging.LogInformation($"Transforming document '{WorkItem.SourceDocId}'")

      If Transformations Is Nothing OrElse Transformations.Count = 0 Then
        If Me.Parent IsNot Nothing Then
          mobjTransformations = Me.Parent.Transformations
        End If
      End If
      Dim lobjDocument As Document = Nothing
      Dim lstrErrorMessage As String = String.Empty

      Dim lobjTransform As Transformation = Nothing
      Dim lstrRootTransformName As String = GetStringParameterValue(PARAM_ROOT_TRANSFORMATION, String.Empty)

      If (String.IsNullOrEmpty(lstrRootTransformName)) Then
        Throw New ArgumentNullException(lstrRootTransformName)
      End If

      lobjTransform = Me.Transformations(lstrRootTransformName)
      If (lobjTransform Is Nothing) Then
        Me.ProcessedMessage = "Unable to transform document, the primary transformation is not initialized."
        menuResult = OperationEnumerations.Result.Failed
        OnError(New OperableErrorEventArgs(Me, WorkItem, Me.ProcessedMessage))
        Return menuResult
      Else
        'RKS - Added this so each batch item has it's own Transform.
        Dim lobjItemTransform = New Transformation(lobjTransform.Actions)
        lobjDocument = lobjItemTransform.TransformDocument(Me.WorkItem.Document, lstrErrorMessage)

        'RKS - original code caused other batch items to impact each other
        'lobjDocument = Me.Transformations.PrimaryTransformation.TransformDocument(Me.WorkItem.Document, lstrErrorMessage)

        If (lobjDocument Is Nothing) Then
          Me.ProcessedMessage = String.Format("{0} - Failed to transform document. Batch item id '{1}' Batch Id '{2}' Error: '{3}'",
                                            Reflection.MethodBase.GetCurrentMethod, Me.DocumentId, Me.Parent.Id, lstrErrorMessage)
          menuResult = OperationEnumerations.Result.Failed
          OnError(New OperableErrorEventArgs(Me, WorkItem, Me.ProcessedMessage))
          Return menuResult
        End If
        Me.WorkItem.Document = lobjDocument
        menuResult = OperationEnumerations.Result.Success

      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, ex))
    End Try

    Return menuResult

  End Function

#End Region

  Private Sub mobjTransformations_CollectionChanged(ByVal sender As Object, ByVal e As System.Collections.Specialized.NotifyCollectionChangedEventArgs) Handles mobjTransformations.CollectionChanged
    Try
      If e.Action = Specialized.NotifyCollectionChangedAction.Add Then
        If Transformations.PrimaryTransformation IsNot Nothing Then
          SetParameterValue(PARAM_ROOT_TRANSFORMATION, Transformations.PrimaryTransformation.Name)
        Else
          SetParameterValue(PARAM_ROOT_TRANSFORMATION, String.Empty)
        End If
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

End Class