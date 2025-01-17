'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Profiles

  <Xml.Serialization.XmlRoot("Profiles")>
  Public Class Profiles
    Inherits CCollection(Of Profile)
    Implements ISerialize

#Region "Class Constants"

    Public Const PROFILES_FILE_EXTENSION As String = "xml"

#End Region

#Region "Class Variables"

    Private ReadOnly mobjProfileConfiguration As ProfileConfiguration

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub


    Public Sub New(ByVal lpProfileConfiguration As ProfileConfiguration)
      Try
        mobjProfileConfiguration = lpProfileConfiguration
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Constructs a Profiles object by deserializing the XML in the specified
    ''' path.
    ''' </summary>
    ''' <param name="lpXMLFilePath">A fully qualified XML file path for the serialized Transformation file.</param>
    Public Sub New(ByVal lpXMLFilePath As String)
      Dim lobjXMLDocument As New Xml.XmlDocument

      Try
        lobjXMLDocument.Load(lpXMLFilePath)
        LoadFromXmlDocument(lobjXMLDocument)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 0, lpXMLFilePath)
        Throw New ApplicationException(String.Format("Unable to create profiles object from xml file: {0}", ex.Message), ex)
      End Try
    End Sub



#End Region
#Region "Public Methods"

    Default Shadows Property Item(ByVal name As String) As Profile
      Get
        Try
          For Each lobjProfile As Profile In Me
            If (lobjProfile.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase)) Then
              Return lobjProfile
            End If
          Next

          Return Nothing

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try

      End Get
      Set(ByVal value As Profile)
        Try
          Dim lobjProfile As Profile
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjProfile = CType(MyBase.Item(lintCounter), Profile)
            If lobjProfile.Name = name Then
              MyBase.Item(lintCounter) = value
              Exit Property
            End If
          Next
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Profiles::Set_Item('{0}')", name))
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Default Shadows Property Item(ByVal Index As Integer) As Profile
      Get
        Try
          Return MyBase.Item(Index)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Profiles::Get_Item('{0}')", Index))
          ' Re-throw the exception to the caller
          Throw
        Finally
        End Try
      End Get
      Set(ByVal value As Profile)
        MyBase.Item(Index) = value
      End Set
    End Property


    ''' <summary>
    ''' Finds a Profile object in the collection by Profile Name
    ''' </summary>
    ''' <param name="lstrProfileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindByProfileName(ByVal lstrProfileName As String) As Profile

      For Each lobjProfile As Profile In Me '.Items
        If (lobjProfile.Name.Equals(lstrProfileName, StringComparison.CurrentCultureIgnoreCase)) Then
          Return lobjProfile
        End If
      Next

      Return Nothing
    End Function

    ''' <summary>
    ''' Finds a Profile object in the collection by Entire Folder Path
    ''' </summary>
    ''' <param name="lstrFolderPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindByMonitoredPath(ByVal lstrFolderPath As String) As Profile

      For Each lobjProfile As Profile In Me '.Items
        'If (lobjProfile.MonitoredLocation.ToLower = lstrFolderPath.ToLower) Then
        '  Return lobjProfile
        'End If
        If (InStr(lstrFolderPath.ToLower, lobjProfile.MonitoredLocation.ToLower) > 0) Then
          Return lobjProfile
        End If
      Next

      Return Nothing

    End Function

    Public Overloads Sub Add(ByVal lpProfile As Profile)
      Try
        lpProfile.SetProfileConfiguration(Me.mobjProfileConfiguration)
        MyBase.Add(lpProfile)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub CheckForDuplicatesOnSave()
      Try
        HandleDuplicatesOnSerialize()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="lpXML"></param>
    ''' <remarks></remarks>
    Private Sub LoadFromXmlDocument(ByVal lpXML As Xml.XmlDocument)
      Try
        Dim lobjProfiles As Profiles = CType(Deserialize(lpXML), Profiles)
        Helper.AssignObjectProperties(lobjProfiles, Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub

    ''' <summary>
    ''' Returns true if there is a profile with the same name 
    ''' </summary>
    ''' <param name="lstrProfileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DoesItemExistMoreThanOnce(ByVal lstrProfileName As String) As Boolean
      Try
        Dim iCount As Integer = 0
        For Each lobjProfile As Profile In Me
          If (lobjProfile.Name.Equals(lstrProfileName, StringComparison.CurrentCultureIgnoreCase)) Then
            iCount += 1
          End If
        Next

        If (iCount > 1) Then Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

      Return False

    End Function

    ''' <summary>
    ''' Returns the first duplicate profile name it finds in the collection
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FindDuplicateProfileName() As String
      Try
        For Each lobjProfile As Profile In Me
          If (DoesItemExistMoreThanOnce(lobjProfile.Name)) Then
            Return lobjProfile.Name
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

      Return String.Empty

    End Function

    Private Sub HandleDuplicatesOnSerialize()
      Try
        Dim lstrDuplicateProfileName As String = FindDuplicateProfileName()
        If (lstrDuplicateProfileName <> String.Empty) Then
          Throw New ApplicationException("A Profile with the name '" & lstrDuplicateProfileName & "' already exists.")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Throw
      End Try
    End Sub

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
        Return PROFILES_FILE_EXTENSION
      End Get
    End Property

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try
        HandleDuplicatesOnSerialize()
        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpFilePath, lpErrorMessage))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function Deserialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try
        HandleDuplicatesOnSerialize()
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        Helper.DumpException(ex)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Function Deserialize(ByVal lpXMLNode As System.Xml.XmlNode) As Profiles
      Try
        Return CType(Serializer.Deserialize.XmlString(lpXMLNode.OuterXml, GetType(Profiles)), Profiles)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try
        HandleDuplicatesOnSerialize()
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      Try
        HandleDuplicatesOnSerialize()
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try
        HandleDuplicatesOnSerialize()
        Serialize(lpFilePath, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize
      Try
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

  End Class

End Namespace

