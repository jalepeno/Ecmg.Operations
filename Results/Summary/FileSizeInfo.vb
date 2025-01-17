' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IProcessingTime.vb
'  Description :  [type_description_here]
'  Created     :  4/29/2016 1:00:40 AM
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

Public Class FileSizeInfo
  Inherits Statistical

#Region "Class Variables"

  Private mobjAverageFileSize As FileSize

#End Region

#Region "Public Properties"

  Public Property AverageFileSize As FileSize
    Get
      Try
        Return mobjAverageFileSize
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As FileSize)
      Try
        mobjAverageFileSize = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

#End Region

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
      'Dim llngTicks As Long = Math.Round(Average * TimeSpan.TicksPerSecond)
      'mobjAverageProcessingTime = New TimeSpan(llngTicks)
      Dim llngAverageBytes As Long = Math.Round(Average)
      mobjAverageFileSize = New FileSize(llngAverageBytes)
      Average = mobjAverageFileSize.ToString()
      Total = New FileSize(CLng(Total)).ToString()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(lpValues As IEnumerable(Of FileSize))
    Try
      GetStatistics(lpValues)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(lpValues As IEnumerable(Of IFileSizeInfo))
    Try
      GetStatistics(lpValues)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

  Public Overloads Sub GetStatistics(lpValues As IEnumerable(Of FileSize))
    Try

      If lpValues Is Nothing Then
        Throw New ArgumentNullException("lpValues")
      End If

      If lpValues.Count = 0 Then
        Throw New ArgumentOutOfRangeException("lpValues", "No values supplied.")
      End If

      Dim _
        lobjTotalFileSize As _
          New FileSize(CLng(Math.Round(Helper.Total(lpValues.Select(Function(x) CDbl(x.Bytes)).ToList(), SampleSize))))
      Total = lobjTotalFileSize
      Maximum = lpValues.Max
      Minimum = lpValues.Min
      mobjAverageFileSize = New FileSize(CLng(Math.Round(lobjTotalFileSize.Bytes / SampleSize)))
      Average = mobjAverageFileSize

      Dim lobjMode As Object

      Dim lobjConvertedValues As New List(Of Double)

      'If mobjAverageFileSize.Gigabytes > 1
      '  lobjConvertedValues = lpValues.Select(Function(x) x.Gigabytes).ToList()
      'ElseIf mobjAverageFileSize.Megabytes > 1
      '  lobjConvertedValues = lpValues.Select(Function(x) x.Megabytes).ToList()
      'ElseIf mobjAverageFileSize.Kilobytes > 1
      '  lobjConvertedValues = lpValues.Select(Function(x) x.Kilobytes).ToList()
      'Else
      '  lobjConvertedValues = lpValues.Select(Function(x) CDbl(x.Bytes)).ToList()
      'End If

      lobjConvertedValues = lpValues.Select(Function(x) CDbl(x.Bytes)).ToList()

      Variance = Helper.Variance(lobjConvertedValues)
      'Variance = Helper.Variance(lpValues)
      StandardDeviation = Helper.StandardDeviation(lobjConvertedValues)
      Median = Helper.Median(lpValues)
      lobjMode = Helper.Mode(lpValues.Select(Function(x) CDbl(x.Bytes)).ToList())
      If lobjMode IsNot Nothing Then
        Mode = New FileSize(CLng(lobjMode))
      End If
      Range = Helper.Range(lpValues)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overloads Sub GetStatistics(lpValues As IEnumerable(Of IFileSizeInfo))
    Try
      If lpValues Is Nothing Then
        Throw New ArgumentNullException("lpValues")
      End If

      If lpValues.Count = 0 Then
        Throw New ArgumentOutOfRangeException("lpValues", "No values supplied.")
      End If

      Dim _
        lobjTotalFileSize As _
          New FileSize(CLng(Math.Round(Helper.Total(lpValues.Select(Function(x) CDbl(x.Total)).ToList(), SampleSize))))
      Total = lobjTotalFileSize
      Maximum = lpValues.Max
      Minimum = lpValues.Min
      mobjAverageFileSize = New FileSize(CLng(Math.Round(lobjTotalFileSize.Bytes / SampleSize)))
      Average = mobjAverageFileSize

      Dim lobjMode As Object

      Dim lobjConvertedValues As New List(Of Double)

      lobjConvertedValues = lpValues.Select(Function(x) CDbl(x.Total)).ToList()

      Variance = Helper.Variance(lobjConvertedValues)
      'Variance = Helper.Variance(lpValues)
      StandardDeviation = Helper.StandardDeviation(lobjConvertedValues)
      'Median = Helper.Median(lpValues)
      lobjMode = Helper.Mode(lpValues.Select(Function(x) CDbl(x.Total)).ToList())
      If lobjMode IsNot Nothing Then
        Mode = New FileSize(CLng(lobjMode))
      End If
      'Range = Helper.Range(lpValues)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overrides Sub ReadXml(reader As XmlReader)
    Try
      With reader
        Total = FileSizeFromAttribute(reader, "total")
        AverageFileSize = FileSizeFromAttribute(reader, "avg")
        Average = AverageFileSize
        Double.TryParse(.GetAttribute("stdev"), StandardDeviation)
        Median = FileSizeFromAttribute(reader, "median")
        Mode = FileSizeFromAttribute(reader, "mode")
        Range = FileSizeFromAttribute(reader, "range")
        Maximum = FileSizeFromAttribute(reader, "max")
        Minimum = FileSizeFromAttribute(reader, "min")
        Variance = .GetAttribute("variance")
        SampleSize = .GetAttribute("sampleSize")
      End With
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Function FileSizeFromAttribute(ByRef lpReader As XmlReader, ByVal lpAttributeLabel As String) As FileSize
    Try
      Dim lstrAttributeValue As String = lpReader.GetAttribute(lpAttributeLabel)
      Dim lobjReturnValue As FileSize

      If Not String.IsNullOrEmpty(lstrAttributeValue) Then
        If IsNumeric(lstrAttributeValue) Then
          Return New FileSize(CLng(lstrAttributeValue))
        Else
          Return FileSize.FromString(lstrAttributeValue)
        End If

      Else
        Return Nothing
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function
End Class

