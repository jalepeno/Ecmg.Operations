'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents
Imports Documents.Configuration
Imports Documents.Core
Imports Documents.Providers
Imports Documents.SerializationUtilities
Imports Documents.Utilities

Namespace Profiles

  <DataContract()>
  <Xml.Serialization.XmlInclude(GetType(KeyValuePairConnectionString))>
  Public Class ProfileConfiguration
    Implements ISerialize
    Implements IXmlSerializable

#Region "Class Variables"

    Private mobjProfiles As New Profiles(Me)
    Private mobjConnectionStrings As New KeyValuePairConnectionStrings
    Private mobjTransformations As New Transformations.TransformationCollection
    Private mobjProcesses As IProcesses = New Processes
    Private mobjSource As ContentSource

#End Region

#Region "Public Properties"

    <DataMember()>
    Public Property ConnectionStrings() As KeyValuePairConnectionStrings
      Get
        Return mobjConnectionStrings
      End Get
      Set(ByVal value As KeyValuePairConnectionStrings)
        mobjConnectionStrings = value
      End Set
    End Property

    <DataMember()>
    Public Property Processes() As IProcesses
      Get
        Return mobjProcesses
      End Get
      Set(ByVal value As IProcesses)
        mobjProcesses = value
      End Set
    End Property

    <DataMember()>
    Public Property Transformations() As Transformations.TransformationCollection
      Get
        Return mobjTransformations
      End Get
      Set(ByVal value As Transformations.TransformationCollection)
        mobjTransformations = value
      End Set
    End Property

    <DataMember()>
    Public Property Profiles() As Profiles
      Get
        Return mobjProfiles
      End Get
      Set(ByVal value As Profiles)
        mobjProfiles = value
      End Set
    End Property

    Public ReadOnly Property SourceContentSource As ContentSource
      Get
        If (mobjSource Is Nothing) Then
          CreateSourceContentSource()
        End If
        Return mobjSource
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpFilePath As String)
      Try

        Dim lobjProfileConfiguration As ProfileConfiguration = CType(Serializer.Deserialize.XmlFile(lpFilePath, GetType(ProfileConfiguration)), ProfileConfiguration)

        Helper.AssignObjectProperties(lobjProfileConfiguration, Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Sub CreateSourceContentSource()

      Try
        mobjSource = CreateFirstContentSource()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub


    Private Shared Function CreateFirstContentSource() As ContentSource
      Try


        Dim lstrDefaultConnectionString As String =
          String.Format("Name=Local File System;Provider=File System Provider;ExportPath={0}Exports\Local File System;ImportPath={0}Exports\Local File System;RootPath=Desktop;UserName=;Password=F8261CB94C60527B;MaxLongFileNameLength=100",
                        FileHelper.Instance.CtsDocsPath)


        Return New ContentSource(lstrDefaultConnectionString)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Methods"

    Public Shared Function CreateFromByteArray(ByVal lpByteArray As Byte()) As ProfileConfiguration

      Try
        Dim lobjMemoryStream As New IO.MemoryStream(lpByteArray)
        Dim lobjProfileConfig As ProfileConfiguration = ProfileConfiguration.CreateFromStream(lobjMemoryStream)
        lobjMemoryStream = Nothing
        Return lobjProfileConfig
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Shared method to create a ProfileConfiguration from a stream
    ''' </summary>
    ''' <param name="lpStream"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateFromStream(ByVal lpStream As IO.Stream) As ProfileConfiguration

      Try

        If (lpStream.CanSeek) Then
          lpStream.Position = 0
        End If

        Return CType(Serializer.Deserialize.FromStream(lpStream, GetType(ProfileConfiguration)), ProfileConfiguration)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Function ToByteArray() As Byte()

      Try
        Dim lobjStream As IO.Stream = Me.ToStream
        Dim lobjByteArray As Byte() = New Byte(Convert.ToInt32(lobjStream.Length)) {}
        lobjStream.Read(lobjByteArray, 0, Convert.ToInt32(lobjStream.Length))
        lobjStream.Close()
        Return lobjByteArray

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Function ToStream() As IO.Stream
      Try
        Return Serializer.Serialize.ToStream(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "ISerialize Implementation"

    ''' <summary>
    ''' Gets the default file extension 
    ''' to be used for serialization 
    ''' and deserialization.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DefaultFileExtension() As String Implements ISerialize.DefaultFileExtension
      Get
        Return "xml"
      End Get
    End Property

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpFilePath, lpErrorMessage))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function Deserialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        Helper.DumpException(ex)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try
        Me.Profiles.CheckForDuplicatesOnSave()
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      Try
        Me.Profiles.CheckForDuplicatesOnSave()
        Serializer.Serialize.XmlFile(Me, lpFilePath)
        Helper.FormatXmlFile(lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try
        Me.Profiles.CheckForDuplicatesOnSave()
        Serialize(lpFilePath, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize
      Try
        Me.Profiles.CheckForDuplicatesOnSave()
        'Serializer.Serialize.XmlFile(Me, lpFilePath, , mstrXMLProcessingInstructions)
        If lpWriteProcessingInstruction = True Then
          Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
        Else
          Serializer.Serialize.XmlFile(Me, lpFilePath)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Function ToString() As String Implements ISerialize.ToXmlString
      Try
        Return Serializer.Serialize.XmlString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IXmlSerializable Implementation"

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      ' As per the Microsoft guidelines this is not implemented
      Return Nothing
    End Function

    Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
      Dim lobjXmlDocument As New XmlDocument
      Dim lobjAttribute As XmlAttribute = Nothing

      Try

        lobjXmlDocument.Load(reader)

        With lobjXmlDocument

          '' Read the ConnectionStrings element
          Dim lobjConnectionStringsNode As XmlNode = .SelectSingleNode("//ConnectionStrings")
          Me.ConnectionStrings.AddRange(GetConnectionStrings(lobjConnectionStringsNode))

          Dim lobjProcessessNode As XmlNode = .SelectSingleNode("//Processes")
          Me.Processes = GetProcesses(lobjProcessessNode)

          Dim lobjProfilesNode As XmlNode = .SelectSingleNode("//Profiles")
          Me.Profiles = GetProfiles(lobjProfilesNode)

          For Each lobjProfile As Profile In Me.Profiles
            lobjProfile.SetProfileConfiguration(Me)
            If (Not String.IsNullOrEmpty(lobjProfile.TransformationPath) AndAlso IO.File.Exists(lobjProfile.TransformationPath)) Then
              Dim lobjTransform As New Transformations.Transformation(lobjProfile.TransformationPath)
              If (Not Me.Transformations.Contains(lobjTransform)) Then
                Me.Transformations.Add(lobjTransform)
              End If
            End If
          Next


        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        lobjXmlDocument = Nothing
      End Try

    End Sub

    Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
      Try

        With writer

          ' Write the ConnectionStrings element
          .WriteStartElement("ConnectionStrings")

          For Each lobjConnectionString As KeyValuePairConnectionString In Me.ConnectionStrings
            .WriteRaw(Serializer.Serialize.XmlElementString(lobjConnectionString))
          Next

          ' Close 'ConnectionStrings' element
          .WriteEndElement()


          ' Write the Processes element
          .WriteStartElement("Processes")

          For Each lobjProcess As IProcess In Me.Processes
            .WriteRaw(lobjProcess.ToXmlElementString)
          Next

          ' Close 'Processes' element
          .WriteEndElement()

          ' Write the Profiles element
          .WriteRaw(Serializer.Serialize.XmlElementString(Me.Profiles))

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Shared Function GetConnectionStrings(ByVal lpConnectionStringsNode As XmlNode) As KeyValuePairConnectionStrings
      Try

        Dim lstrConnectionString As String = String.Empty
        Dim lobjConnectionStringValueNode As XmlNode = Nothing
        Dim lobjKeyValuePairConnectionString As KeyValuePairConnectionString = Nothing
        Dim lobjKeyValuePairConnectionStrings As New KeyValuePairConnectionStrings

        ' Read the ConnectionStrings element
        For Each lobjConnectionString As XmlNode In lpConnectionStringsNode.ChildNodes

          lobjConnectionStringValueNode = lobjConnectionString.SelectSingleNode("Value")
          If lobjConnectionStringValueNode IsNot Nothing Then
            lstrConnectionString = lobjConnectionStringValueNode.InnerText
            lobjKeyValuePairConnectionString = New KeyValuePairConnectionString(lstrConnectionString)
            lobjKeyValuePairConnectionStrings.Add(lobjKeyValuePairConnectionString)
          End If

        Next

        Return lobjKeyValuePairConnectionStrings

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Shared Function GetProcesses(ByVal lpProcessessNode As XmlNode) As IProcesses
      Dim lobjProcesses As New Processes

      Try

        For Each lobjProcessNode As XmlNode In lpProcessessNode.ChildNodes
          Dim lobjProcess As IProcess = Process.Deserialize(lobjProcessNode)
          If lobjProcess IsNot Nothing Then
            lobjProcesses.Add(lobjProcess)
          End If
        Next

        Return lobjProcesses

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Shared Function GetProfiles(ByVal lpProfilesNode As XmlNode) As Profiles
      Try

        Dim lobjProfiles As Profiles = Profiles.Deserialize(lpProfilesNode)

        Return lobjProfiles

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
