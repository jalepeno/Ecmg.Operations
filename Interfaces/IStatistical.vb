' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IStatistical.vb
'  Description :  [type_description_here]
'  Created     :  4/29/2016 12:46:40 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Public Interface IStatistical
  Property Total As Object

  Property Maximum As Object
  Property Minimum As Object

  Property Average As Object

  ''' <summary>
  ''' Variance is the measure of the amount of variation of all the scores (not just the extremes which give the range).
  ''' </summary>
  ''' <returns></returns>
  Property Variance As Object

  ''' <summary>
  ''' The standard deviation of processing times.
  ''' </summary>
  ''' <returns></returns>
  Property StandardDeviation As Double

  ''' <summary>
  ''' Median is the number separating the higher half of the processing times from the lower half.
  ''' </summary>
  ''' <returns></returns>
  Property Median As Object

  ''' <summary>
  ''' Mode is the most frequently occuring processing time.
  ''' </summary>
  ''' <returns></returns>
  Property Mode As Object

  Property Range As Object
  Property SampleSize As Integer

  Function ToJsonString() As String
  Function ToXmlString() As String
  Function ToXmlElementString() As String
End Interface
