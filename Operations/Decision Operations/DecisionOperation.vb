' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  DecisionOperation.vb
'  Description :  [type_description_here]
'  Created     :  4/13/2012 11:00:35 PM
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

Public MustInherit Class DecisionOperation
  Inherits Operation
  Implements IDecisionOperation

#Region "Class Constants"

  Friend Const PARAM_VERSION_SCOPE As String = "VersionScope"
  Friend Const PARAM_CONTENT_ELEMENT_INDEX As String = "ContentElementIndex"
  Friend Const PARAM_MODE As String = "Mode"
  Friend Const PARAM_EVALUATION_ACTION As String = "EvaluationAction"

#End Region

#Region "Class Enumerations"

  Public Enum ModeEnum
    Valid
    Invalid
  End Enum

  ''' <summary>Specifies the behavior following the evaluation.</summary>
  Public Enum EvaluationActionEnum
    ''' <summary>The default behavior will execute any available true operations if the evaluation returns true or any false operations if the evalation returns false.</summary>
    [Default]
    ''' <summary>Will execute any true operations if the evaluation returns true or will fail the operation if the evaluation returns false.</summary>
    FailOnFalse
    ''' <summary>Will execute any false operations if the evaluation returns false or will fail the operation if the evaluation returns true.</summary>
    FailOnTrue
  End Enum

#End Region

#Region "Class Variables"

  Private mobjTrueOperations As New Operations
  Private mobjFalseOperations As New Operations
  Private mblnEvaluation As Nullable(Of Boolean)
  'Private mstrDescription As String = String.Empty

#End Region

  Public Overrides ReadOnly Property CanRollback As Boolean
    Get
      Return True
    End Get
  End Property

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try
      Return ExecuteDecision()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Overrides Function OnRollback() As Result
    Try
      mblnEvaluation = Evaluate()

      If mblnEvaluation = True Then
        Return TrueOperations.Rollback(Me.WorkItem)
      Else
        Return FalseOperations.Rollback(Me.WorkItem)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Sub Reset()
    Try
      MyBase.Reset()
      For Each lobjOperableStep As IOperable In TrueOperations
        lobjOperableStep.Reset()
      Next
      For Each lobjOperableStep As IOperable In FalseOperations
        lobjOperableStep.Reset()
      Next
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Function ExecuteDecision() As OperationEnumerations.Result
    Try
      mblnEvaluation = Evaluate()

      Dim lenuEvaluationAction As EvaluationActionEnum = GetEnumParameterValue(PARAM_EVALUATION_ACTION,
            GetType(EvaluationActionEnum), EvaluationActionEnum.Default)


      Select Case lenuEvaluationAction
        Case EvaluationActionEnum.Default
          If mblnEvaluation = True Then
            Return TrueOperations.Execute(Me.WorkItem)
          Else
            Return FalseOperations.Execute(Me.WorkItem)
          End If

        Case EvaluationActionEnum.FailOnFalse
          If mblnEvaluation = True Then
            Return TrueOperations.Execute(Me.WorkItem)
          Else
            With Me.WorkItem
              .ProcessedMessage = String.Format("{0} failed on false decision.", Me.Name)
              .ProcessedStatus = ProcessedStatus.Failed
              Return OperationEnumerations.Result.Failed
            End With
          End If

        Case EvaluationActionEnum.FailOnTrue
          If mblnEvaluation = True Then
            With Me.WorkItem
              .ProcessedMessage = String.Format("{0} failed on true decision.", Me.Name)
              .ProcessedStatus = ProcessedStatus.Failed
              Return OperationEnumerations.Result.Failed
            End With
          Else
            Return FalseOperations.Execute(Me.WorkItem)
          End If

        Case Else
          If mblnEvaluation = True Then
            Return TrueOperations.Execute(Me.WorkItem)
          Else
            Return FalseOperations.Execute(Me.WorkItem)
          End If

      End Select

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#Region "Protected Methods"

  Public Overrides Sub CheckParameters()
    Try
      UpdateParameterToEnum(PARAM_VERSION_SCOPE, GetType(VersionScopeEnum))
      UpdateParameterToEnum(PARAM_MODE, GetType(ModeEnum))
      UpdateParameterToEnum(PARAM_EVALUATION_ACTION, GetType(EvaluationActionEnum))
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
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_VERSION_SCOPE,
          VersionScopeEnum.AllVersions, GetType(VersionScopeEnum),
          "Specifies which version(s) to use for the evaluation."))
      End If

      If lobjParameters.Contains(PARAM_CONTENT_ELEMENT_INDEX) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmLong, PARAM_CONTENT_ELEMENT_INDEX, 0,
          "Specifies which content element to use for the evaluation.  NOTE: The first content element is specified with a value of zero."))
      End If

      If lobjParameters.Contains(PARAM_MODE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_MODE, ModeEnum.Valid, GetType(ModeEnum),
          "Specifies whether the mode is valid or invalid.  When the mode is valid, the evaluation will return true if the extension exists, when the mode is invalid, the evaluation will return false if the extension exists."))
      End If

      If lobjParameters.Contains(PARAM_EVALUATION_ACTION) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_EVALUATION_ACTION, EvaluationActionEnum.Default, GetType(EvaluationActionEnum),
          "Specifies the behavior following the evaluation."))
      End If
      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "IDecisionOperation Implementation"

  Public MustOverride Function Evaluate() As Boolean Implements IDecisionOperation.Evaluate

  Public ReadOnly Property Evaluation As Boolean Implements IDecisionOperation.Evaluation
    Get
      Try
        Return CBool(mblnEvaluation)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Property FalseOperations As IOperations Implements IDecisionOperation.FalseOperations
    Get
      Try
        Return mobjFalseOperations
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperations)
      Try
        mobjFalseOperations = CType(value, Operations)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public ReadOnly Property RunOperations As IOperations Implements IDecisionOperation.RunOperations
    Get
      Try
        If mblnEvaluation.HasValue = False Then
          Return New Operations
        Else
          If Evaluation = True Then
            Return TrueOperations
          Else
            Return FalseOperations
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Property TrueOperations As IOperations Implements IDecisionOperation.TrueOperations
    Get
      Try
        Return mobjTrueOperations
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As IOperations)
      Try
        mobjTrueOperations = CType(value, Operations)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

#End Region

  Public Overrides Sub WriteXml(ByVal writer As System.Xml.XmlWriter)
    Try
      MyBase.WriteXml(writer)

      With writer

        ' Open the TrueOperations Element
        .WriteStartElement("TrueOperations")

        If TrueOperations IsNot Nothing AndAlso TrueOperations.Count > 0 Then
          For Each lobjOperation As IOperable In TrueOperations
            ' Write the Parameter element
            .WriteRaw(lobjOperation.ToXmlElementString)
          Next
        End If

        ' End the TrueOperations element
        .WriteEndElement()

        ' Open the FalseOperations Element
        .WriteStartElement("FalseOperations")

        If FalseOperations IsNot Nothing AndAlso FalseOperations.Count > 0 Then
          For Each lobjOperation As IOperable In FalseOperations
            ' Write the Parameter element
            .WriteRaw(lobjOperation.ToXmlElementString)
          Next
        End If

        ' End the FalseOperations element
        .WriteEndElement()

      End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

End Class
