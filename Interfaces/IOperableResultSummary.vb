' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IOperableResultSummary.vb
'  Description :  [type_description_here]
'  Created     :  4/26/2016 10:22:40 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Public Interface IOperableResultSummary

  Property Name As String

  ''' <summary>
  ''' Property CreateDate As DateTime
  ''' </summary>
  ''' <remarks>
  ''' Added on 9/11/2024 to mark when the summary was created.  
  ''' The primary reason is to be able to determine whether or 
  ''' not to update depending on how stale it is.
  ''' </remarks>
  ''' <returns></returns>
  Property CreateDate As DateTime

  Property Node As String
  Property Result As OperationEnumerations.Result
  Property Scope As OperationScope

  Property ProcessingTime As ProcessingTimeInfo

  'Property AverageProcessingTime As TimeSpan
  'Dim ldblVariance As Double
  'Dim ldblStandardDev As Double
  'Dim ldblMedian As Double
  'Dim ldblMode As Double
  'Dim ldblRange As Double


  '''' <summary>
  '''' Variance is the measure of the amount of variation of all the scores (not just the extremes which give the range).
  '''' </summary>
  '''' <returns></returns>
  'Property Variance As Double

  '''' <summary>
  '''' The standard deviation of processing times.
  '''' </summary>
  '''' <returns></returns>
  'Property StandardDeviation As Double

  '''' <summary>
  '''' Median is the number separating the higher half of the processing times from the lower half.
  '''' </summary>
  '''' <returns></returns>
  'Property Median As Double

  '''' <summary>
  '''' Mode is the most frequently occuring processing time.
  '''' </summary>
  '''' <returns></returns>
  'Property Mode As Nullable(Of Double)
  'Property Range As Double
  'Property SampleSize As Integer
  ReadOnly Property Parent As IOperable

  Function ToJsonString() As String
  Function ToXmlString() As String
  Function ToXmlElementString() As String

End Interface
