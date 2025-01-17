' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ProcessingTimeInfo.vb
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

Public Class ProcessingTimeInfo
  Inherits Statistical

#Region "Class Variables"

  Private mobjAverageProcessingTime As TimeSpan

#End Region

#Region "Public Properties"

  Public Property AverageProcessingTime As TimeSpan
    Get
      Try
        Return mobjAverageProcessingTime
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As TimeSpan)
      Try
        mobjAverageProcessingTime = value
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
  ''' Constructs a new ProcessingTime instance using a collection of seconds.
  ''' </summary>
  ''' <param name="lpValues"></param>
  Public Sub New(lpValues As IEnumerable(Of Double))
    Try
      GetStatistics(lpValues)
      Dim llngTicks As Long = Math.Round(Average * TimeSpan.TicksPerSecond)
      mobjAverageProcessingTime = New TimeSpan(llngTicks)
      Average = mobjAverageProcessingTime.ToString()
      Total = New TimeSpan(CLng(Math.Round(Total * TimeSpan.TicksPerSecond))).ToString()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(lpValues As IEnumerable(Of TimeSpan))
    Try
      GetStatistics(lpValues)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(lpValue As TimeSpan)
    Try
      Dim lobjValues As New List(Of TimeSpan) From {
        lpValue
      }

      GetStatistics(lobjValues)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

  Protected Friend Overrides Function DebuggerIdentifier() As String
    Dim lobjIdentifierBuilder As New Text.StringBuilder
    Try

      If Total IsNot Nothing Then
        lobjIdentifierBuilder.AppendFormat("Total: {0}", FriendlyTimeSpanString(Total))
        lobjIdentifierBuilder.AppendFormat(" / Min: {0}", FriendlyTimeSpanString(Minimum))
        lobjIdentifierBuilder.AppendFormat(" / Max: {0}", FriendlyTimeSpanString(Maximum))
        lobjIdentifierBuilder.AppendFormat(" / Avg: {0}", FriendlyTimeSpanString(Average))
        lobjIdentifierBuilder.AppendFormat(" / StDev: {0}", StandardDeviation)
        lobjIdentifierBuilder.AppendFormat(" / Range: {0}", FriendlyTimeSpanString(Range))
      Else
        lobjIdentifierBuilder.Append("Not Initialized")
      End If

      Return lobjIdentifierBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Return lobjIdentifierBuilder.ToString
    End Try
  End Function

  Public Overloads Sub GetStatistics(lpValues As IEnumerable(Of TimeSpan))
    Try

#If NET8_0_OR_GREATER Then
      ArgumentNullException.ThrowIfNull(lpValues)
#Else
          If lpValues Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpValues))
          End If
#End If
      If Not lpValues.Any() Then
        Throw New ArgumentOutOfRangeException(NameOf(lpValues), "No values supplied.")
      End If

      'Dim lobjTotalTimeSpan As _
      '    New TimeSpan(CLng(Math.Round(Helper.Total(lpValues.Select(Function(x) CDbl(x.Ticks)).ToList(), SampleSize))))

      'Dim ldblTicksCalculation As Double = Math.Round(Helper.Total(lpValues.Select(Function(x) CDbl(x.Ticks)).ToList(), SampleSize))
      Dim ldblTicksCalculation As Double = Math.Round(Helper.Total(lpValues.Select(Function(x) CDbl(x.Ticks)).ToList()))
      Dim llngTicks As Long

      Dim v = Long.TryParse(ldblTicksCalculation, llngTicks)

      Dim lobjTotalTimeSpan As New TimeSpan(llngTicks)

      Total = lobjTotalTimeSpan
      Maximum = lpValues.Max
      Minimum = lpValues.Min
      mobjAverageProcessingTime = New TimeSpan(CLng(Math.Round(lobjTotalTimeSpan.Ticks / lpValues.Count)))
      Average = mobjAverageProcessingTime

      Dim lobjMode As Object

      Dim lobjConvertedValues As New List(Of Double)

      If mobjAverageProcessingTime.Seconds < 1 Then
        lobjConvertedValues = lpValues.Select(Function(x) x.TotalSeconds).ToList()
      ElseIf mobjAverageProcessingTime.Minutes < 1 Then
        lobjConvertedValues = lpValues.Select(Function(x) x.TotalMinutes).ToList()
      Else
        lobjConvertedValues = lpValues.Select(Function(x) x.TotalHours).ToList()
      End If

      Variance = Helper.Variance(lobjConvertedValues)
      StandardDeviation = Helper.StandardDeviation(lobjConvertedValues)
      Median = Helper.Median(lpValues)
      lobjMode = Helper.Mode(lpValues)
      If lobjMode IsNot Nothing Then
        Mode = lobjMode
      End If
      Range = Helper.Range(lpValues)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overrides Sub ReadXml(reader As XmlReader)
    Try
      With reader
        Total = TimeSpanFromAttribute(reader, "total")
        AverageProcessingTime = TimeSpanFromAttribute(reader, "avg")
        Average = AverageProcessingTime
        Dim v = Double.TryParse(.GetAttribute("stdev"), StandardDeviation)
        Median = TimeSpanFromAttribute(reader, "median")
        Mode = TimeSpanFromAttribute(reader, "mode")
        Range = TimeSpanFromAttribute(reader, "range")
        Maximum = TimeSpanFromAttribute(reader, "max")
        Minimum = TimeSpanFromAttribute(reader, "min")
        Variance = .GetAttribute("variance")
        SampleSize = .GetAttribute("sampleSize")
      End With
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Shared Function TimeSpanFromAttribute(ByRef lpReader As XmlReader, ByVal lpAttributeLabel As String) As TimeSpan
    Try
      Dim lstrAttributeValue As String = lpReader.GetAttribute(lpAttributeLabel)
      Dim lobjReturnValue As TimeSpan

      If Not String.IsNullOrEmpty(lstrAttributeValue) Then
        If TimeSpan.TryParse(lstrAttributeValue, lobjReturnValue) Then
          Return lobjReturnValue
        Else
          Return Nothing
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

  Private Shared Function FriendlyTimeSpanString(lpTimeSpan As TimeSpan) As String
    Try
      If lpTimeSpan.TotalSeconds < 1 Then
        Return String.Format("{0} Msecs", lpTimeSpan.TotalMilliseconds)
      ElseIf lpTimeSpan.TotalMinutes < 1 Then
        Return String.Format("{0} Secs", lpTimeSpan.Seconds)
      ElseIf lpTimeSpan.TotalHours < 1 Then
        Return String.Format("{0}:{1} Mins", lpTimeSpan.Minutes, lpTimeSpan.Seconds)
      ElseIf lpTimeSpan.TotalDays < 1 Then
        Return String.Format("{0}:{1}:{2} Hours", lpTimeSpan.Hours, lpTimeSpan.Minutes, lpTimeSpan.Seconds)
      Else
        Return String.Format("{0}:{1}:{2}:{3} Days", lpTimeSpan.Days, lpTimeSpan.Hours, lpTimeSpan.Minutes, lpTimeSpan.Seconds)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

End Class
