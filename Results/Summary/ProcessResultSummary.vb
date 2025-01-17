' ---------------------------------------------------------------------------------
'  Document    :  ProcessResultSummary.vb
'  Description :  [type_description_here]
'  Created     :  4/26/2016 5:28:40 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports System.Xml.Serialization
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Public Class ProcessResultSummary
  Inherits OperationResultSummary
  Implements IProcessResultSummary
  Implements IXmlSerializable

#Region "Class Variables"

  Private mobjFileCountInfo As New FileCountInfo
  Private mobjFileSizeInfo As New FileSizeInfo
  Private mobjVersionCountInfo As New VersionCountInfo

  Private mobjOperationResults As IOperableResultsSummary = New OperationResultsSummary

#End Region

#Region "Public Properties"

  Public Property VersionCountInfo As IStatistical Implements IProcessResultSummary.VersionCountInfo
    Get
      Try
        Return mobjVersionCountInfo
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IStatistical)
      Try
        mobjVersionCountInfo = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property FileCountInfo As IStatistical Implements IProcessResultSummary.FileCountInfo
    Get
      Try
        Return mobjFileCountInfo
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IStatistical)
      Try
        mobjFileCountInfo = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property FileSizeInfo As IStatistical Implements IProcessResultSummary.FileSizeInfo
    Get
      Try
        Return mobjFileSizeInfo
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IStatistical)
      Try
        mobjFileSizeInfo = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property OperationResults As IOperableResultsSummary Implements IProcessResultSummary.OperationResults
    Get
      Return mobjOperationResults
    End Get
    Set(ByVal value As IOperableResultsSummary)
      mobjOperationResults = value
    End Set
  End Property

#End Region


#Region "Constructors"

  Public Sub New()
  End Sub

  Public Sub New(ByVal lpProcessResults As IProcessResults)
    Try
      AssignFromProcessResults(lpProcessResults)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(ByVal lpProcessResultSummaries As IProcessResultSummaries)
    Try
      AssignFromProcessResultSummaries(lpProcessResultSummaries)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Methods"

  Public Shared Function FromXmlString(lpXml As String) As IProcessResultSummary
    Try
      Return Serializer.Deserialize.XmlString(lpXml.Replace("''", "'"), GetType(ProcessResultSummary))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Sub AssignFromProcessResultSummaries(ByVal lpProcessResultSummaries As IProcessResultSummaries)
    Try

      Dim lobjTotalProcessingTimeValueObjects As IEnumerable(Of Object)
      Dim lobjTotalFileSizeTimeValueObjects As IEnumerable(Of IStatistical)
      Dim lobjTotalProcessingTimeValues As New List(Of TimeSpan)
      Dim lobjTotalFileSizeValues As New List(Of FileSizeInfo)
      Dim lobjTotalFileCountValues As New List(Of Double)
      Dim lobjTotalVersionCountValues As New List(Of Double)

      Dim lobjSummaryResults As New List(Of OperationEnumerations.Result)

      lobjTotalProcessingTimeValueObjects = lpProcessResultSummaries.Select(Function(v) v.ProcessingTime.Total)

      For Each lobjTotal As Object In lobjTotalProcessingTimeValueObjects
        If TypeOf lobjTotal Is TimeSpan Then
          lobjTotalProcessingTimeValues.Add(DirectCast(lobjTotal, TimeSpan))
        End If
      Next

      lobjTotalProcessingTimeValues.AddRange(lobjTotalProcessingTimeValues)

      lobjTotalFileSizeTimeValueObjects = From v In lpProcessResultSummaries Where v.FileSizeInfo IsNot Nothing Select v.FileSizeInfo
      For Each lobjFileSize As IStatistical In lobjTotalFileSizeTimeValueObjects
        lobjTotalFileSizeValues.Add(DirectCast(lobjFileSize, FileSizeInfo))
      Next

      'lobjTotalFileSizeValues.AddRange(From v In lpProcessResultSummaries Where v.FileSizeInfo IsNot Nothing Select v.FileSizeInfo)


      lobjTotalFileCountValues.AddRange(lpProcessResultSummaries.Select(Function(v) CDbl(v.FileCountInfo.Total)))
      lobjTotalVersionCountValues.AddRange(lpProcessResultSummaries.Select(Function(v) CDbl(v.VersionCountInfo.Total)))

      For Each lobjProcessResultSummary In lpProcessResultSummaries
        lobjSummaryResults.Add(lobjProcessResultSummary.Result)
      Next

      Me.ProcessingTime = New ProcessingTimeInfo(lobjTotalProcessingTimeValues)
      If lobjTotalFileSizeValues.Count > 0 Then
        '  Me.FileSizeInfo = New FileSizeInfo(lobjTotalFileSizeValues)
      End If
      Me.FileCountInfo = New FileCountInfo(lobjTotalFileCountValues)
      Me.VersionCountInfo = New VersionCountInfo(lobjTotalVersionCountValues)

      With lpProcessResultSummaries.FirstOrDefault()
        Me.Name = .Name
        For Each lobjOperationResult As IOperableResult In .OperationResults
          Me.OperationResults.Add(New OperationResultSummary(lobjOperationResult))
        Next
      End With

      'For Each lobjOperationResult As IOperableResultSummary In Me.OperationResults
      '  lobjOperationResultGroup = lpProcessResults.GetOperationResultsByName(lobjOperationResult.Name)

      '  lobjTotalProcessingTimeValues.Clear()
      '  Try
      '    lobjTotalProcessingTimeValues.AddRange(
      '    lobjOperationResultGroup.Select(Function(v) v.TotalProcessingTime))
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    ' Just keep going
      '  End Try

      'lobjOperationResult.ProcessingTime = New ProcessingTimeInfo(lobjTotalProcessingTimeValues)

      'Next

      For Each lobjResult As OperationEnumerations.Result In lobjSummaryResults
        If lobjResult <> Result.Success Then
          Me.Result = Result.Failed
        End If
      Next
      Me.Result = Result.Success

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub AssignFromProcessResults(ByVal lpProcessResults As IProcessResults)
    Try

      'SetParent(lpProcess)

      If lpProcessResults Is Nothing Then
        Throw New ArgumentNullException("lpProcessResults")
      End If

      If lpProcessResults.Count = 0 Then
        If Not Helper.CallStackContainsMethodName("RunPreOperationChecks") Then
          Throw New ArgumentOutOfRangeException("lpProcessResults", "Process results have no values.")
        Else
          ApplicationLogging.WriteLogEntry("Unable to assign process result summary.  No process results available.")
          Exit Sub
        End If

      End If

      'Dim lobjTotalProcessingTimeValues As New List(Of Double)
      Dim lobjTotalProcessingTimeValues As New List(Of TimeSpan)
      Dim lobjTotalFileSizeValues As New List(Of FileSize)
      Dim lobjTotalContentCountValues As New List(Of Double)
      Dim lobjTotalVersionCountValues As New List(Of Double)

      ''Me.SampleSize = lpProcessResults.Count
      'Dim average = New TimeSpan(lpProcessResults.Select(Function(ts) ts.Ticks).Average())
      'Dim lintAverageTicks =
      '      Aggregate processResult In lpProcessResults Into Average(processResult.TotalProcessingTime.Ticks)
      'Me.AverageProcessingTime = New TimeSpan(lintAverageTicks)

      lobjTotalProcessingTimeValues.AddRange(
        lpProcessResults.Select(Function(v) v.TotalProcessingTime))

      ' Where(TotalContentSize IsNot Nothing)
      'lobjTotalFileSizeValues.AddRange(lpProcessResults.Select(Function(v) v.TotalContentSize))
      lobjTotalFileSizeValues.AddRange(From v In lpProcessResults Where v.TotalContentSize IsNot Nothing Select v.TotalContentSize)
      lobjTotalContentCountValues.AddRange(lpProcessResults.Select(Function(v) CDbl(v.ContentCount)))
      lobjTotalVersionCountValues.AddRange(lpProcessResults.Select(Function(v) CDbl(v.VersionCount)))

      Me.ProcessingTime = New ProcessingTimeInfo(lobjTotalProcessingTimeValues)
      If lobjTotalFileSizeValues.Count > 0 Then
        Me.FileSizeInfo = New FileSizeInfo(lobjTotalFileSizeValues)
      End If
      Me.FileCountInfo = New FileCountInfo(lobjTotalContentCountValues)
      Me.VersionCountInfo = New VersionCountInfo(lobjTotalVersionCountValues)

      'GetStatistics(Me, lobjTotalProcessingTimeValues)

      With lpProcessResults.FirstOrDefault()
        Me.Name = .Name
        Me.Node = .Node
        Me.Result = .Result
        For Each lobjOperationResult As IOperableResult In .OperationResults
          Me.OperationResults.Add(New OperationResultSummary(lobjOperationResult))
        Next
      End With

      'Dim lobjOperableResults As IOperableResults = From operableResult In lpProcessResults Group By 
      Dim lobjOperationResultGroup As IOperableResults


      For Each lobjOperationResult As IOperableResultSummary In Me.OperationResults
        lobjOperationResultGroup = lpProcessResults.GetOperationResultsByName(lobjOperationResult.Name)
        'lobjOperationResult.AverageProcessingTime =
        '  New TimeSpan(
        '    Aggregate processResult In lobjOperationResultGroup Into Average(processResult.TotalProcessingTime.Ticks))

        lobjTotalProcessingTimeValues.Clear()
        Try
          lobjTotalProcessingTimeValues.AddRange(
          lobjOperationResultGroup.Select(Function(v) v.TotalProcessingTime))
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Just keep going
        End Try

        'lobjOperationResult.SampleSize = lobjTotalProcessingTimeValues.Count

        lobjOperationResult.ProcessingTime = New ProcessingTimeInfo(lobjTotalProcessingTimeValues)
        'ldblVariance  = Helper.Variance(lobjTotalProcessingTimeValues)
        'ldblStandardDev = Helper.StandardDeviation(lobjTotalProcessingTimeValues)
        'ldblMedian = Helper.Median(lobjTotalProcessingTimeValues)
        'ldblMode = Helper.Mode(lobjTotalProcessingTimeValues)
        'ldblRange = Helper.Range(lobjTotalProcessingTimeValues)

        'GetStatistics(lobjOperationResult, lobjTotalProcessingTimeValues)
        'ldblCovariance = Helper.Covariance(lobjTotalProcessingTimeValues)

      Next

      'Me.OperationResults.GetItemByName()

      'For Each lobjProcessResult As IProcessResult In lpProcessResults       
      '  'Me.OperationResults.Add(New OperationResultSummary(lobjProcessResult))
      'Next

      '    '' Assign the results from all of the event operations
      '    'SetEventResults(lpProcess.RunBeforeJobBegin, Me.RunBeforeJobBeginResults)
      '    'SetEventResults(lpProcess.RunAfterJobComplete, Me.RunAfterJobCompleteResults)
      '    'SetEventResults(lpProcess.RunOnJobFailure, Me.RunOnJobFailureResults)
      '    'SetEventResults(lpProcess.RunBeforeParentBegin, Me.RunBeforeParentBeginResults)
      '    'SetEventResults(lpProcess.RunAfterParentComplete, Me.RunAfterParentCompleteResults)
      '    'SetEventResults(lpProcess.RunOnParentFailure, Me.RunOnParentFailureResults)
      '    'SetEventResults(lpProcess.RunBeforeBegin, Me.RunBeforeBeginResults)
      '    'SetEventResults(lpProcess.RunAfterComplete, Me.RunAfterCompleteResults)
      '    'SetEventResults(lpProcess.RunOnFailure, Me.RunOnFailureResults)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  'Private Sub GetStatistics(ByRef lpSummary As IOperableResultSummary, lpValues As IEnumerable(Of Double))
  '  Try

  '    If lpSummary Is Nothing
  '      Throw New ArgumentNullException("lpSummary")
  '    End If

  '    If lpValues Is Nothing
  '      Throw New ArgumentNullException("lpValues")
  '    End If

  '    If lpValues.Count = 0
  '      Throw New ArgumentOutOfRangeException("lpValues", "No values supplied.")
  '    End If

  '    With lpSummary
  '      .Variance = Helper.Variance(lpValues)
  '      .StandardDeviation = Helper.StandardDeviation(lpValues)
  '      .Median = Helper.Median(lpValues)
  '      .Mode = Helper.Mode(lpValues)
  '      .Range = Helper.Range(lpValues)
  '    End With

  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
  '    '  Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub

  Protected Friend Overrides Function DebuggerIdentifier() As String
    Return MyBase.DebuggerIdentifier()
  End Function
  '  Private Sub SetEventResults(ByRef lpEventOperable As IOperable, ByRef lpEventOperableResults As IOperableResultsSummary)
  '  Try

  '    If lpEventOperable IsNot Nothing Then
  '      lpEventOperableResults = New OperationResultsSummary
  '      If TypeOf (lpEventOperable) Is IOperation Then
  '        lpEventOperableResults.Add(New OperationResultSummary(lpEventOperable))
  '      ElseIf TypeOf (lpEventOperable) Is IProcess Then
  '        For Each lobjOperation As IOperable In CType(lpEventOperable, IProcess).Operations
  '          lpEventOperableResults.Add(New OperationResultSummary(lpEventOperable))
  '        Next
  '      End If
  '    End If

  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub

#End Region
End Class
