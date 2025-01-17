' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Statistical.vb
'  Description :  [type_description_here]
'  Created     :  4/30/2016 10:49:40 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.Serialization
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region


<DebuggerDisplay("{DebuggerIdentifier(),nq}")>
Public MustInherit Class Statistical
  Implements IStatistical
  Implements IXmlSerializable

#Region "Class Variables"

  Private mobjTotal As Object = 0
  Private mobjMaximum As Object = 0
  Private mobjMinimum As Object = 0
  Private mobjAverage As Object = 0
  Private mobjMedian As Object
  Private mobjMode As Object
  Private mobjRange As Object
  Private mintSampleSize As Integer
  Private mdblStandardDeviation As Double
  Private mobjVariance As Object

#End Region

#Region "Public Properties"

  Public Property Average As Object Implements IStatistical.Average
    Get
      Try
        Return mobjAverage
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Object)
      Try
        mobjAverage = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Median As Object Implements IStatistical.Median
    Get
      Try
        Return mobjMedian
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Object)
      Try
        mobjMedian = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Mode As Object Implements IStatistical.Mode
    Get
      Try
        Return mobjMode
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Object)
      Try
        mobjMode = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Range As Object Implements IStatistical.Range
    Get
      Try
        Return mobjRange
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Object)
      Try
        mobjRange = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property SampleSize As Integer Implements IStatistical.SampleSize
    Get
      Try
        Return mintSampleSize
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Integer)
      Try
        mintSampleSize = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property StandardDeviation As Double Implements IStatistical.StandardDeviation
    Get
      Try
        Return mdblStandardDeviation
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Double)
      Try
        mdblStandardDeviation = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Total As Object Implements IStatistical.Total
    Get
      Try
        Return mobjTotal
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Object)
      Try
        mobjTotal = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Variance As Object Implements IStatistical.Variance
    Get
      Try
        Return mobjVariance
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Object)
      Try
        mobjVariance = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Maximum As Object Implements IStatistical.Maximum
    Get
      Try
        Return mobjMaximum
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Object)
      Try
        mobjMaximum = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Minimum As Object Implements IStatistical.Minimum
    Get
      Try
        Return mobjMinimum
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Object)
      Try
        mobjMinimum = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

#End Region

#Region "Public Methods"

  Public Overridable Sub GetStatistics(lpValues As IEnumerable(Of Double))
    Try

      If lpValues IsNot Nothing Then
        If Not lpValues.Any() Then
          Throw New ArgumentOutOfRangeException(NameOf(lpValues), "No values supplied.")
        End If
        Maximum = lpValues.Max
        Minimum = lpValues.Min
        Total = Helper.Total(lpValues, SampleSize)
        Average = Total / SampleSize
        Variance = Helper.Variance(lpValues)
        StandardDeviation = Helper.StandardDeviation(lpValues)
        ' StandardDeviation = Helper.StDev(lpValues)
        Median = Helper.Median(lpValues)
        Dim lobjMode As Object = Helper.Mode(lpValues)
        If lobjMode IsNot Nothing Then
          Mode = lobjMode
        End If

        Range = Helper.Range(lpValues)
      Else
        Throw New ArgumentNullException(NameOf(lpValues))
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Function ToJsonString() As String Implements IStatistical.ToJsonString
    Try
      'Return Serializer.Serialize.JsonString(Me)
      ' We will do manual JSON serialization for complete speed and control

      Dim lobjStringBuilder As New StringBuilder
      Dim lobjStringWriter As New StringWriter(lobjStringBuilder)

      Using lobjJSONWriter As New JsonTextWriter(lobjStringWriter)
        With lobjJSONWriter
          .Formatting = Newtonsoft.Json.Formatting.Indented

          .WriteRaw(String.Format("{""{0}"": ", Me.GetType.Name))

          .WriteStartObject()

          .WritePropertyName("total")
          .WriteValue(Total)

          .WritePropertyName("avg")
          .WriteValue(Average)

          .WritePropertyName("stdev")
          .WriteValue(StandardDeviation)

          .WritePropertyName("median")
          .WriteValue(Median)

          .WritePropertyName("mode")
          .WriteValue(Mode)

          .WritePropertyName("range")
          .WriteValue(Range)

          .WritePropertyName("variance")
          .WriteValue(Variance)

          .WritePropertyName("sampleSize")
          .WriteValue(SampleSize)

          .WriteEndObject()

          .WriteRaw("}")

        End With
      End Using

      Return lobjStringBuilder.ToString

    Catch Ex As Exception
      ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function ToXmlString() As String Implements IStatistical.ToXmlString
    Try
      Return Serializer.Serialize.XmlString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function ToXmlElementString() As String Implements IStatistical.ToXmlElementString
    Try
      Return Serializer.Serialize.XmlElementString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

#Region "IXmlSerializable Implementation"

  Public Function GetSchema() As XmlSchema Implements IXmlSerializable.GetSchema
    Try
      Return Nothing
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overridable Sub ReadXml(reader As XmlReader) Implements IXmlSerializable.ReadXml
    Try
      With reader
        Total = .GetAttribute("total")
        Average = .GetAttribute("avg")
        Dim v = Double.TryParse(.GetAttribute("stdev"), StandardDeviation)
        Median = .GetAttribute("median")
        Mode = .GetAttribute("mode")
        Range = .GetAttribute("range")
        Maximum = .GetAttribute("max")
        Minimum = .GetAttribute("min")
        Variance = .GetAttribute("variance")
        Dim v1 = Integer.TryParse(.GetAttribute("sampleSize"), SampleSize)
      End With
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overridable Sub WriteXml(writer As XmlWriter) Implements IXmlSerializable.WriteXml
    Try
      With writer

        If Total IsNot Nothing Then
          .WriteAttributeString("total", Total.ToString())
        Else
          .WriteAttributeString("total", String.Empty)
        End If

        If Minimum IsNot Nothing Then
          .WriteAttributeString("min", Minimum.ToString())
        Else
          .WriteAttributeString("min", String.Empty)
        End If

        If Maximum IsNot Nothing Then
          .WriteAttributeString("max", Maximum.ToString())
        Else
          .WriteAttributeString("max", String.Empty)
        End If

        If Average IsNot Nothing Then
          .WriteAttributeString("avg", Average.ToString())
        Else
          .WriteAttributeString("avg", String.Empty)
        End If

        .WriteAttributeString("stdev", StandardDeviation.ToString())

        If Median IsNot Nothing Then
          .WriteAttributeString("median", Median.ToString())
        Else
          .WriteAttributeString("median", String.Empty)
        End If

        If Mode IsNot Nothing Then
          .WriteAttributeString("mode", Mode.ToString())
        Else
          .WriteAttributeString("mode", String.Empty)
        End If

        If Range IsNot Nothing Then
          .WriteAttributeString("range", Range.ToString())
        Else
          .WriteAttributeString("range", String.Empty)
        End If

        If Variance IsNot Nothing Then
          .WriteAttributeString("variance", Variance.ToString())
        Else
          .WriteAttributeString("variance", String.Empty)
        End If

        .WriteAttributeString("sampleSize", SampleSize.ToString())

      End With
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#End Region

#Region "Protected Methods"

  Protected Friend Overridable Function DebuggerIdentifier() As String
    Dim lobjIdentifierBuilder As New Text.StringBuilder
    Try

      If Total IsNot Nothing Then
        lobjIdentifierBuilder.AppendFormat("Total: {0:N3}", Total.ToString)
        lobjIdentifierBuilder.AppendFormat(" / Min: {0:N3}", Minimum)
        lobjIdentifierBuilder.AppendFormat(" / Max: {0:N3}", Maximum)
        lobjIdentifierBuilder.AppendFormat(" / Avg: {0:N3}", Average)
        lobjIdentifierBuilder.AppendFormat(" / StDev: {0:N3}", StandardDeviation)
        lobjIdentifierBuilder.AppendFormat(" / Range: {0:N3}", Range)
      Else
        lobjIdentifierBuilder.Append("Not Initialized")
      End If

      Return lobjIdentifierBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Return lobjIdentifierBuilder.ToString
    End Try
  End Function

#End Region
End Class
