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
Imports Documents
Imports Documents.Utilities

#End Region

Namespace Profiles

  <DataContract()>
  Public Class RecordProfile
    Implements ICloneable
    Implements INotifyPropertyChanged


#Region "Class Variables"

    Private mobjTransformation As Transformations.Transformation
    Private mstrTransformationPath As String = ""
    Private mobjContentSource As Providers.ContentSource
    Private mstrDestinationContentSource As String = ""
    'Private mobjLocation As Ecmg.Cts.Records.Location
    Private mstrRecordFolder As String = ""
    Private mstrRecordClass As String = ""
    Private mobjDefaultDocumentProperties As New Core.ECMProperties
    Private mobjDefaultVersionProperties As New Core.ECMProperties
    Private mobjProfile As Profile 'Parent object

#End Region

#Region "Constructors"

    Public Sub New()
    End Sub

    Public Sub New(ByVal lpProfile As Profile)
      mobjProfile = lpProfile
    End Sub

#End Region

#Region "Public Properties"
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
        mstrDestinationContentSource = value
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
    Public Property RecordFolder() As String
      Get
        Return mstrRecordFolder
      End Get
      Set(ByVal value As String)
        mstrRecordFolder = value
      End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property RecordClass() As String
      Get
        Return mstrRecordClass
      End Get
      Set(ByVal value As String)
        mstrRecordClass = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the default document scoped properties 
    ''' </summary>
    ''' <value>An ECMProperties collection object</value>
    ''' <returns>An ECMProperties collection object</returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property DocumentProperties() As Core.ECMProperties
      Get
        Return mobjDefaultDocumentProperties
      End Get
      Set(ByVal value As Core.ECMProperties)
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
    Public Property VersionProperties() As Core.ECMProperties
      Get
        Return mobjDefaultVersionProperties
      End Get
      Set(ByVal value As Core.ECMProperties)
        mobjDefaultVersionProperties = value
      End Set
    End Property

    '''' <summary>
    '''' Gets or sets the destination folder into which documents should be filed
    '''' </summary>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks>Specified using the correct fully qualified syntax for the destination repository</remarks>
    '<DataMember()> _
    'Public Property Location() As Ecmg.Cts.Records.Location
    '  Get
    '    Return mobjLocation
    '  End Get
    '  Set(ByVal value As Ecmg.Cts.Records.Location)
    '    mobjLocation = value
    '  End Set
    'End Property

#End Region

#Region "Private Methods"

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

          If (mobjProfile.ProfileConfiguration IsNot Nothing) Then

            mobjContentSource = New Providers.ContentSource(mobjProfile.ProfileConfiguration.ConnectionStrings(lpName).Value)

          Else
            'Throw an exception
            Throw New Exception(String.Format("Unable to create record content source in Profile '{0}' ProfileConfiguration is nothing, cannot continue. Content source name is '{1}'", mobjProfile.Name, lpName))
          End If

        Else
          'Throw an exception
          Throw New Exception(String.Format("Unable to create record content source in Profile '{0}' Content Source Name is nothing, cannot continue.", mobjProfile.Name))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Error connecting to content source: " + lpName, TraceEventType.Error)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Friend Methods"

    Friend Sub SetProfile(ByVal lpProfile As Profile)
      Try
        mobjProfile = lpProfile
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "ICloneable"
    Public Function Clone() As Object Implements System.ICloneable.Clone
      Dim lobjProfile As New RecordProfile()
      Try

        lobjProfile.ContentSourceName = Me.ContentSourceName
        lobjProfile.RecordClass = Me.RecordClass
        lobjProfile.RecordFolder = Me.RecordFolder
        lobjProfile.TransformationPath = Me.TransformationPath
        If (Me.DocumentProperties IsNot Nothing) Then
          lobjProfile.DocumentProperties = CType(Me.DocumentProperties.Clone, Core.ECMProperties)
        End If
        If (Me.VersionProperties IsNot Nothing) Then
          lobjProfile.VersionProperties = CType(Me.VersionProperties.Clone, Core.ECMProperties)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

      Return lobjProfile
    End Function
#End Region

#Region "INotifyPropertyChanged"

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Protected Sub OnPropertyChanged(ByVal name As String)
      RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub

#End Region

  End Class

End Namespace