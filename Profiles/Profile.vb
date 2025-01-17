'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Text
Imports Documents
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Operations.PlugIns

#End Region

Namespace Profiles

#Region "Public Enumerations"

  ''' <summary>
  ''' Designates the type of location to be monitored
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum LocationType
    NOTDEFINED = 0
    ''' <summary>
    ''' A Windows local or network drive
    ''' </summary>
    ''' <remarks></remarks>
    NTFS = 1
    ''' <summary>
    ''' An email address
    ''' </summary>
    ''' <remarks></remarks>
    SMTP = 2
    ''' <summary>
    ''' An ftp site
    ''' </summary>
    ''' <remarks>Should include the complete ftp folder path</remarks>
    FTP = 3
  End Enum

#End Region

  ''' <summary>
  ''' Represents a single collector profile
  ''' </summary>
  ''' <remarks></remarks>
  <DataContract(), KnownType(GetType(PlugInProperty))>
  Public Class Profile
    Implements INotifyPropertyChanged
    Implements ISerialize
    Implements ICloneable
    Implements IComparable
    Implements IItemParent


#Region "Class Constants"

    Public Const PROFILE_FILE_EXTENSION As String = "xml"

#End Region

#Region "Class Variables"

    Private mstrName As String
    Private mstrMonitoredLocation As String
    Private mblnIsMonitoredLocationCDF As Boolean
    Private menuLocationType As LocationType = LocationType.NTFS
    Private mblnIncludeSubFolders As Boolean
    Private mstrMonitoredFileFilter As String = "*.*"
    Private mstrDestinationContentSource As String = ""
    Private mobjContentSource As Providers.ContentSource
    Private mobjTransformation As Transformations.Transformation
    Private mstrTransformationPath As String = ""
    Private mstrFolderToFileIn As String = ""
    Private mblnAppendSubFolderToFolderFiledIn As Boolean
    Private mobjDefaultDocumentProperties As New PlugInProperties
    Private mobjDefaultVersionProperties As New PlugInProperties
    Private mobjScanInterval As Integer = 5000
    Private mobjRecordProfile As RecordProfile
    Private mstrDocumentClass As String
    Private mobjPlugIns As New PlugInConfigurations
    Private mobjProfileConfiguration As ProfileConfiguration 'Parent object that holds a collection of connection strings and transforms
    Private mobjProcess As IProcess
    Private mstrProcessName As String = ""
#End Region

#Region "Public Properties"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property RecordProfile() As RecordProfile
      Get
        Return mobjRecordProfile
      End Get
      Set(ByVal value As RecordProfile)
        mobjRecordProfile = value
        mobjRecordProfile?.SetProfile(Me)
        OnPropertyChanged("RecordProfile")
      End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property DocumentClass() As String
      Get
        Return mstrDocumentClass
      End Get
      Set(ByVal value As String)
        mstrDocumentClass = value
      End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property ScanInterval() As Integer
      Get
        Return mobjScanInterval
      End Get
      Set(ByVal value As Integer)
        mobjScanInterval = value
      End Set
    End Property

    ''' <summary>
    ''' Which file type to monitor
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property MonitoredFileFilter() As String
      Get
        Return mstrMonitoredFileFilter
      End Get
      Set(ByVal value As String)
        mstrMonitoredFileFilter = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the mblnIsMonitoredLocationCDF property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property IsMonitoredLocationCDF() As Boolean
      Get
        Return mblnIsMonitoredLocationCDF
      End Get
      Set(ByVal value As Boolean)
        mblnIsMonitoredLocationCDF = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the IncludeSubFolders property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property IncludeSubFolders() As Boolean
      Get
        Return mblnIncludeSubFolders
      End Get
      Set(ByVal value As Boolean)
        mblnIncludeSubFolders = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the name of the profile
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>It is useful to provide a name for all profiles so that they 
    ''' may be more easily identified within the collection of profiles</remarks>
    <Xml.Serialization.XmlAttribute()>
    <DataMember()>
    Public Property Name() As String Implements IItemParent.Name
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    ''' <summary>
    ''' The location to be monitored for this profile
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property MonitoredLocation() As String
      Get
        Return mstrMonitoredLocation
      End Get
      Set(ByVal value As String)
        mstrMonitoredLocation = value
        OnPropertyChanged("MonitoredLocation")
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the location type for this profile
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property LocationType() As LocationType
      Get
        Return menuLocationType
      End Get
      Set(ByVal value As LocationType)
        menuLocationType = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the destination content source for the documents to be loaded into
    ''' </summary>
    ''' <value>The name of the content source</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property ContentSourceName() As String
      Get
        Return mstrDestinationContentSource
      End Get
      Set(ByVal value As String)
        If (value <> mstrDestinationContentSource) Then
          mobjContentSource = Nothing
        End If
        'Dim prevValue As String = mstrDestinationContentSource
        mstrDestinationContentSource = value
        'If (prevValue <> mstrDestinationContentSource) Then
        'OnPropertyChanged("ContentSourceName")
        'End If
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the Process Name
    ''' </summary>
    ''' <value>The name of the content source</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property ProcessName() As String
      Get
        Return mstrProcessName
      End Get
      Set(ByVal value As String)
        mstrProcessName = value
      End Set
    End Property

    ''' <summary>
    ''' The destination content source for incoming documents
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ContentSource() As Providers.ContentSource
      Get
        If (mobjContentSource Is Nothing) Then
          CreateContentSource(Me.mstrDestinationContentSource)
        End If
        Return mobjContentSource
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets the path to a transformation file
    ''' </summary>
    ''' <value>A fully qualified path for a transformation file</value>
    ''' <returns>A transformation file path</returns>
    ''' <remarks>If the path is a valid transformation path this will 
    ''' also attempt to set the object reference for the 
    ''' Transformation property.</remarks>
    <DataMember()>
    Public Property TransformationPath() As String
      Get
        Return mstrTransformationPath
      End Get
      Set(ByVal value As String)
        mstrTransformationPath = value
        OnPropertyChanged("TransformationPath")
      End Set
    End Property

    Public ReadOnly Property Transformation() As Transformations.Transformation
      Get
        If (mobjTransformation Is Nothing) Then
          CreateTransformation(mstrTransformationPath)
        End If
        Return mobjTransformation
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets the destination folder into which documents should be filed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Specified using the correct fully qualified syntax for the destination repository</remarks>
    <DataMember()>
    Public Property FolderToFileIn() As String
      Get
        Return mstrFolderToFileIn
      End Get
      Set(ByVal value As String)
        mstrFolderToFileIn = value
        OnPropertyChanged("FolderToFileIn")
      End Set
    End Property

    ''' <summary>
    ''' Append the entire sub-folder path to FolderFiledIn property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property AppendSubFolderToFolderFiledIn() As Boolean
      Get
        Return mblnAppendSubFolderToFolderFiledIn
      End Get
      Set(ByVal value As Boolean)
        mblnAppendSubFolderToFolderFiledIn = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the default document scoped properties 
    ''' </summary>
    ''' <value>An ECMProperties collection object</value>
    ''' <returns>An ECMProperties collection object</returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property DocumentProperties() As PlugInProperties
      Get
        Return mobjDefaultDocumentProperties
      End Get
      Set(ByVal value As PlugInProperties)
        mobjDefaultDocumentProperties = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the default version scoped properties 
    ''' </summary>
    ''' <value>An ECMProperties collection object</value>
    ''' <returns>An ECMProperties collection object</returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property VersionProperties() As PlugInProperties
      Get
        Return mobjDefaultVersionProperties
      End Get
      Set(ByVal value As PlugInProperties)
        mobjDefaultVersionProperties = value
      End Set
    End Property

    ''' <summary>
    ''' Collection of plugins to use for this profile
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property PlugIns() As PlugInConfigurations
      Get
        Return mobjPlugIns
      End Get
      Set(ByVal value As PlugInConfigurations)
        mobjPlugIns = value
      End Set
    End Property

    ''' <summary>
    ''' Parent object that contains a collection of connection strings and transformations
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ProfileConfiguration() As ProfileConfiguration
      Get
        Return mobjProfileConfiguration
      End Get
    End Property

    ''' <summary>
    ''' Process
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public ReadOnly Property Process() As IProcess
      Get
        If (mobjProcess Is Nothing) Then
          CreateProcess(mstrProcessName)
        End If
        Return mobjProcess
      End Get
    End Property

    ''' <summary>
    ''' Transformations
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Xml.Serialization.XmlIgnore()>
    Public Property Transformations As Transformations.TransformationCollection Implements IItemParent.Transformations
      Get
        Return mobjProfileConfiguration.Transformations
      End Get
      Set(value As Transformations.TransformationCollection)
        mobjProfileConfiguration.Transformations = value
      End Set
    End Property


    <Xml.Serialization.XmlIgnore()>
    Public Property DestinationConnection As Core.IRepositoryConnection Implements IItemParent.DestinationConnection
      Get
        Try
          Return ContentSource
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Core.IRepositoryConnection)
        Try
          If TypeOf value Is Providers.ContentSource Then
            mobjContentSource = CType(value, Providers.ContentSource)
          Else
            ApplicationLogging.WriteLogEntry(
              String.Format("Unable to assign DestinationConnection from value, value '{0}' is not a valid ContentSource object.",
                            value.ConnectionString), Reflection.MethodBase.GetCurrentMethod, TraceEventType.Error, 73512)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <Xml.Serialization.XmlIgnore()>
    Public Property ExportPath As String Implements IItemParent.ExportPath
      Get
        Return String.Empty
      End Get
      Set(value As String)

      End Set
    End Property

    <Xml.Serialization.XmlIgnore()>
    Public Property Id As String Implements IItemParent.Id
      Get
        Return String.Empty
      End Get
      Set(value As String)

      End Set
    End Property

    <Xml.Serialization.XmlIgnore()>
    Public Property SourceConnection As Core.IRepositoryConnection Implements IItemParent.SourceConnection
      Get
        Return mobjProfileConfiguration.SourceContentSource
      End Get
      Set(value As Core.IRepositoryConnection)

      End Set
    End Property

    Public Sub RefreshDestinationConnection() Implements IItemParent.RefreshDestinationConnection
      Try
        ' This method is to satisfy the IItemParent interface, but for now we will not attempt to recreate the ContentSource, if needed we can change this in the future.
        ' Do Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub RefreshSourceConnection() Implements IItemParent.RefreshSourceConnection
      Try
        ' This method is to satisfy the IItemParent interface, but for now we will not attempt to recreate the ContentSource, if needed we can change this in the future.
        ' Do Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Constructors"

    Public Sub New()
    End Sub

    Public Sub New(ByVal lpProfileName As String, ByVal lpContentSourceName As String)
      Me.New(lpProfileName, lpContentSourceName, "", "")
    End Sub

    Public Sub New(ByVal lpProfileName As String,
                   ByVal lpContentSourceName As String,
                   ByVal lpTransformationPath As String,
                   ByVal lpMonitoredLocation As String)
      Try
        mstrName = lpProfileName
        ContentSourceName = lpContentSourceName
        'mobjContentSource = New Providers.ContentSource(ConnectionStrings(lpContentSourceName).ConnectionString)
        TransformationPath = lpTransformationPath
        If (lpTransformationPath.Length > 0) Then
          mobjTransformation = New Transformations.Transformation(lpTransformationPath)
        End If
        mstrMonitoredLocation = lpMonitoredLocation
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpProfileName As String,
                   ByVal lpContentSourceName As String,
                   ByVal lpTransformationPath As String,
                   ByVal lpMonitoredLocation As String,
                   ByVal lpDocumentProperties As PlugInProperties,
                   ByVal lpVersionProperties As PlugInProperties)

      Me.New(lpProfileName, lpContentSourceName, lpTransformationPath, lpMonitoredLocation)

      Try
        DocumentProperties = lpDocumentProperties
        VersionProperties = lpVersionProperties
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpProfileName As String,
                   ByVal lpContentSourceName As String,
                   ByVal lpTransformationPath As String,
                   ByVal lpMonitoredLocation As String,
                   ByVal lpFolderToFileIn As String,
                   ByVal lpDocumentProperties As PlugInProperties,
                   ByVal lpVersionProperties As PlugInProperties)
      Try
        mstrName = lpProfileName
        ContentSourceName = lpContentSourceName
        'mobjContentSource = New Providers.ContentSource(ConnectionStrings(lpContentSourceName).ConnectionString)
        TransformationPath = lpTransformationPath
        If (lpTransformationPath.Length > 0) Then
          mobjTransformation = New Transformations.Transformation(lpTransformationPath)
        End If
        mstrFolderToFileIn = lpFolderToFileIn
        mstrMonitoredLocation = lpMonitoredLocation
        DocumentProperties = lpDocumentProperties
        VersionProperties = lpVersionProperties
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Sub InitializeRecordProfile()
      Try
        Me.RecordProfile = New RecordProfile(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Friend Methods"

    Friend Sub SetProfileConfiguration(ByVal lpProfileConfiguration As ProfileConfiguration)
      Try
        mobjProfileConfiguration = lpProfileConfiguration
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Sub CreateProcess(ByVal lpName As String)

      Try

        For Each lobjProcess In mobjProfileConfiguration.Processes

          If (lobjProcess.Name = lpName) Then
            mobjProcess = lobjProcess
            Exit Sub
          End If

        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    Private Sub CreateTransformation(ByVal lpPath As String)
      Try

        If (lpPath <> String.Empty) Then
          ' Only set the transformation object if the path is valid
          If File.Exists(lpPath) Then
            mobjTransformation = New Transformations.Transformation(lpPath)
          Else
            Throw New Exception(String.Format(
                                 "Unable to create transformation object from TransformationPath '{0}'.  The path is not valid.",
                                 lpPath))
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Given a connection string name, creates the initialized content source (mobjContentSource)
    ''' </summary>
    ''' <param name="lpName"></param>
    ''' <remarks></remarks>
    Private Sub CreateContentSource(ByVal lpName As String)

      Try
        If (lpName <> String.Empty) Then

          If (mobjProfileConfiguration IsNot Nothing) Then

            mobjContentSource = New Providers.ContentSource(mobjProfileConfiguration.ConnectionStrings(lpName).Value)

          Else
            'Throw an exception
            Throw New Exception(String.Format("Unable to create content source in Profile '{0}' ProfileConfiguration is nothing, cannot continue. Content source name is '{1}'", Me.Name, lpName))
          End If

        Else
          'Throw an exception
          Throw New Exception(String.Format("Unable to create content source in Profile '{0}' Content Source Name is nothing, cannot continue.", Me.Name))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Error connecting to content source: " + lpName, TraceEventType.Error)
        '  Re-throw the exception to the caller
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
        Return PROFILE_FILE_EXTENSION
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
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      Try
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try
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

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Try
        Return Serializer.Serialize.XmlString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Function ToString() As String
      Try
        Dim lobjStringBuilder As New StringBuilder

        lobjStringBuilder.Append(Me.Name)

        If Not String.IsNullOrEmpty(Me.ProcessName) Then
          lobjStringBuilder.AppendFormat(": {0}", Me.ProcessName)
        End If

        Return lobjStringBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "INotifyPropertyChanged"

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Protected Sub OnPropertyChanged(ByVal name As String)
      RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub

#End Region

#Region "ICloneable"
    Public Function Clone() As Object Implements System.ICloneable.Clone
      Dim lobjProfile As New Profile()
      Try

        lobjProfile.Name = Me.Name
        lobjProfile.ContentSourceName = Me.ContentSourceName
        lobjProfile.DocumentClass = Me.DocumentClass
        If (Me.DocumentProperties IsNot Nothing) Then
          lobjProfile.DocumentProperties = CType(Me.DocumentProperties.Clone, PlugInProperties)
        End If
        lobjProfile.FolderToFileIn = Me.FolderToFileIn
        lobjProfile.IncludeSubFolders = Me.IncludeSubFolders
        lobjProfile.IsMonitoredLocationCDF = Me.IsMonitoredLocationCDF
        lobjProfile.LocationType = Me.LocationType
        lobjProfile.MonitoredFileFilter = Me.MonitoredFileFilter
        lobjProfile.MonitoredLocation = Me.MonitoredLocation
        lobjProfile.PlugIns = Me.PlugIns
        If (Me.RecordProfile IsNot Nothing) Then
          lobjProfile.RecordProfile = CType(Me.RecordProfile.Clone, RecordProfile)
        End If
        lobjProfile.ScanInterval = Me.ScanInterval
        lobjProfile.TransformationPath = Me.TransformationPath
        If (Me.VersionProperties IsNot Nothing) Then
          lobjProfile.VersionProperties = CType(Me.VersionProperties.Clone, PlugInProperties)
        End If
        lobjProfile.SetProfileConfiguration(Me.ProfileConfiguration)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

      Return lobjProfile
    End Function
#End Region

#Region "IComparable"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
      Try
        Dim lobjProfile As Profile = CType(obj, Profile)
        Return String.Compare(Me.Name, lobjProfile.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region


  End Class

End Namespace
