' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  DoubleStatistical.vb
'  Description :  [type_description_here]
'  Created     :  5/5/2016 10:35:40 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports System.Xml
Imports Documents.Utilities

#End Region

Public MustInherit Class DoubleStatistical
  Inherits Statistical

#Region "Constructors"

  Public Sub New()

  End Sub

  ''' <summary>
  ''' Constructs a new FileSizeInfo instance using a collection of bytes.
  ''' </summary>
  ''' <param name="lpValues"></param>
  Public Sub New(lpValues As IEnumerable(Of Double))
    Try
      GetStatistics(lpValues)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

  Public Overrides Sub ReadXml(reader As XmlReader)
    Try
      With reader
        'Total = .GetAttribute("total")
        'Average = .GetAttribute("avg")
        'StandardDeviation = .GetAttribute("stdev")
        'Median = .GetAttribute("median")
        'Mode = .GetAttribute("mode")
        'Range = .GetAttribute("range")
        'Maximum = .GetAttribute("max")
        'Minimum = .GetAttribute("min")
        'Variance = .GetAttribute("variance")
        'SampleSize = .GetAttribute("sampleSize")
        Double.TryParse(.GetAttribute("total"), Total)
        Double.TryParse(.GetAttribute("avg"), Average)
        Double.TryParse(.GetAttribute("stdev"), StandardDeviation)
        Double.TryParse(.GetAttribute("median"), Median)
        Double.TryParse(.GetAttribute("mode"), Mode)
        Double.TryParse(.GetAttribute("range"), Range)
        Double.TryParse(.GetAttribute("max"), Maximum)
        Double.TryParse(.GetAttribute("min"), Minimum)
        Double.TryParse(.GetAttribute("variance"), Variance)
        Integer.TryParse(.GetAttribute("sampleSize"), SampleSize)
      End With
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

End Class
