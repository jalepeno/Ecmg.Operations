'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents
Imports Documents.Arguments
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.SerializationUtilities
Imports Documents.Utilities


#End Region

Namespace PlugIns

  ''' <summary>
  ''' Base Class for all PlugIns
  ''' </summary>
  ''' <remarks></remarks>
  <Serializable()>
  <Xml.Serialization.XmlInclude(GetType(PlugInPropertyDefinition))>
  <Xml.Serialization.XmlInclude(GetType(PlugInProperty))>
  <Xml.Serialization.XmlInclude(GetType(ContentSource))>
  <Xml.Serialization.XmlInclude(GetType(Transformations.Transformation))>
  Public MustInherit Class CPlugIn
    Implements PlugIns.IPlugIn
    Implements IDisposable

#Region "Private Properties"
    Private WithEvents MobjPlugInPropertyDefinitions As New PlugInPropertyDefinitions
    Private mobjPlugInProperties As New PlugInProperties
    Private mstrPlugInPath As String = String.Empty
    Private mstrPlugInClassName As String = String.Empty
    Private mstrVersion As String = String.Empty
    Private menumPlugInType As PlugInType
    Private menumPlugInExecuteTiming As PlugInExecuteTiming
    Private menumPlugInExecutionMode As PlugInExecutionMode
#End Region

#Region "Public Properties"

    Public MustOverride ReadOnly Property Name() As String Implements IPlugIn.Name
    Public MustOverride ReadOnly Property Version() As String Implements IPlugIn.Version
    ''' <summary>Describes the function of this plugin</summary>
    Public MustOverride ReadOnly Property Description() As String Implements IPlugIn.Description
    Public MustOverride ReadOnly Property PlugInClassName() As String Implements IPlugIn.PlugInClassName
    Public MustOverride ReadOnly Property PlugInExecutionMode() As PlugInExecutionMode Implements IPlugIn.PlugInExecutionMode
    '''<summary>Tell the caller when to Execute this plugin</summary>
    Public MustOverride ReadOnly Property PlugInExecuteTiming() As PlugInExecuteTiming Implements IPlugIn.PlugInExecuteTiming
    Public MustOverride ReadOnly Property PlugInType() As PlugInType Implements IPlugIn.PlugInType

    ''' <summary>
    ''' Gets the collection of plugin property definitions
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Xml.Serialization.XmlIgnore()>
    Public ReadOnly Property PlugInPropertyDefinitions() As PlugInPropertyDefinitions Implements IPlugIn.PlugInPropertyDefinitions
      Get
        Return MobjPlugInPropertyDefinitions
      End Get
    End Property

    ''' <summary>
    ''' Gets/Sets the collection of plugin properties
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PlugInProperties() As PlugInProperties Implements IPlugIn.PlugInProperties
      Get
        Return mobjPlugInProperties
      End Get
      Set(ByVal value As PlugInProperties)
        mobjPlugInProperties = value
      End Set
    End Property

    ''' <summary>
    ''' File path to the plugin
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PlugInPath() As String
      Get
        Return mstrPlugInPath
      End Get
      Set(ByVal value As String)
        mstrPlugInPath = value
      End Set
    End Property

#End Region

#Region "Private Properties"

    Private ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
      Try
        Initialize()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub
#End Region

#Region "Private Methods"
    ''' <summary>
    ''' Initializes the plugin
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Initialize()
      Try
        AddProperties()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Throw
      End Try
    End Sub

    Private Sub AddProperties()
      ' Add the base property definitions here.
      Try
        'Add in any base property definitions here
        'mobjPlugInPropertyDefinitions.Add(New PlugInPropertyDefinition("MyProperty", GetType(System.String), True))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

    End Sub
#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Main Execute method
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Function Execute() As PlugInExecuteReturnArgs Implements IPlugIn.Execute

    ''' <summary>
    ''' Factory method for instantiating a plugin from its dll path
    ''' </summary>
    ''' <param name="lpPlugInPath">Fully qualified file name of plugin file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Create(ByVal lpPlugInPath As String) As CPlugIn

      Try
        Dim lobjAssembly As System.Reflection.Assembly
        Dim lobjPlugInCandidate As Type
        Dim lobjPlugIn As CPlugIn


        ' Check to make sure that the ProviderPath has a value, if not we will throw an exception
        If lpPlugInPath Is Nothing Then
          Throw New InvalidPathException("Unable to get plugin, lpPlugInPath is Nothing", "")
        End If
        If lpPlugInPath.Length = 0 Then
          Throw New InvalidPathException("Unable to get plugin, lpPlugInPath is a zero length string", "")
        End If

        ' Check to make sure the path is valid, if not we will throw an exception
        If IO.File.Exists(lpPlugInPath) = False Then
          Dim ioex As New IO.FileNotFoundException(String.Format("File does not exist: '{0}'.", lpPlugInPath), lpPlugInPath)
          Throw New InvalidPathException(String.Format("The plugin path '{0}' does not point to a valid file.", lpPlugInPath), lpPlugInPath, ioex)
        End If

        lobjAssembly = Reflection.Assembly.LoadFrom(lpPlugInPath)

        For Each lobjType As Type In lobjAssembly.GetTypes
          lobjPlugInCandidate = lobjType.GetInterface("IPlugIn")
          If lobjPlugInCandidate IsNot Nothing Then
            lobjPlugIn = CType(lobjAssembly.CreateInstance(lobjType.FullName), CPlugIn)
            If (lobjPlugIn IsNot Nothing) Then
              lobjPlugIn.PlugInPath = lpPlugInPath
              lobjPlugIn.PlugInProperties = lobjPlugIn.PlugInPropertyDefinitions.ToPlugInProperties()
            End If
            Return lobjPlugIn
          End If
        Next

        Return Nothing

      Catch TargetEx As Reflection.TargetInvocationException
        If TargetEx.InnerException IsNot Nothing Then
          Throw TargetEx.InnerException
        Else
          'ApplicationLogging.LogException(TargetEx, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End If
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Factory method for instantiating a plugin from its PlugInConfiguration
    ''' </summary>
    ''' <param name="lpPlugInConfiguration">PlugInConfiguration</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Create(ByVal lpPlugInConfiguration As PlugInConfiguration) As CPlugIn

      Try

        Dim lobjPlugIn As CPlugIn = Create(lpPlugInConfiguration.PlugInPath)

        If (lobjPlugIn IsNot Nothing) Then

          'Set the properties of the plugin
          'lobjPlugIn.SetPlugInClassName(lpPlugInConfiguration.PlugInClassName)
          'lobjPlugIn.SetPlugInExecuteTiming(lpPlugInConfiguration.PlugInExecuteTiming)
          'lobjPlugIn.SetPlugInExecutionMode(lpPlugInConfiguration.PlugInExecutionMode)
          lobjPlugIn.PlugInPath = lpPlugInConfiguration.PlugInPath
          lobjPlugIn.PlugInProperties = lpPlugInConfiguration.PlugInProperties
          'lobjPlugIn.SetPlugInType(lpPlugInConfiguration.PlugInType)
          'lobjPlugIn.SetVersion(lpPlugInConfiguration.Version)

          Return lobjPlugIn
        End If

        Return Nothing

      Catch TargetEx As Reflection.TargetInvocationException
        If TargetEx.InnerException IsNot Nothing Then
          Throw TargetEx.InnerException
        Else
          'ApplicationLogging.LogException(TargetEx, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End If
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Function ToPlugInConfiguration() As PlugInConfiguration

      Dim lobjPIC As New PlugInConfiguration

      Try
        lobjPIC.Description = Me.Description
        lobjPIC.Name = Me.Name
        'lobjPIC.PlugInClassName = Me.PlugInClassName
        'lobjPIC.PlugInExecuteTiming = Me.PlugInExecuteTiming
        'lobjPIC.PlugInExecutionMode = Me.PlugInExecutionMode
        lobjPIC.PlugInPath = Me.PlugInPath
        lobjPIC.PlugInProperties = Me.PlugInProperties
        'lobjPIC.PlugInType = Me.PlugInType
        'lobjPIC.Version = Me.Version
        Return lobjPIC
      Catch ex As Exception
        Throw
      End Try

    End Function

    Sub Serialize(ByVal lpFilePath As String)
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function Serialize() As System.Xml.XmlDocument
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Dim lobjPlugIn As CPlugIn = CType(Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType), CPlugIn)
        Return lobjPlugIn
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpFilePath, lpErrorMessage))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function DeSerialize(ByVal lpXML As System.Xml.XmlDocument) As Object
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        'Helper.DumpException(ex)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Sub SetVersion(ByVal lpValue As String)
      mstrVersion = lpValue
    End Sub

    Protected Sub SetPlugInType(ByVal lpValue As PlugInType)
      menumPlugInType = lpValue
    End Sub

    Protected Sub SetPlugInClassName(ByVal lpValue As String)
      mstrPlugInClassName = lpValue
    End Sub

    Protected Sub SetPlugInExecuteTiming(ByVal lpValue As PlugInExecuteTiming)
      menumPlugInExecuteTiming = lpValue
    End Sub

    Protected Sub SetPlugInExecutionMode(ByVal lpValue As PlugInExecutionMode)
      menumPlugInExecutionMode = lpValue
    End Sub

#End Region

#Region "Events"
    Public Event ExecuteBegin(ByVal lpMessage As String) Implements IPlugIn.ExecuteBegin
    Public Event ExecuteComplete(ByVal lpMessage As String) Implements IPlugIn.ExecuteComplete
    Public Event ExecuteError(ByVal lpMessage As String) Implements IPlugIn.ExecuteError
    Public Event ExecuteReportProgress(ByVal lpPercentProgress As Integer, ByVal lpMessage As String) Implements IPlugIn.ExecuteReportProgress


    ''' <summary>
    ''' Wrapper for ExecuteBegin Event
    ''' </summary>
    ''' <param name="lpMessage"></param>
    ''' <remarks></remarks>
    Protected Sub RaiseExecuteBegin(ByVal lpMessage As String)
      RaiseEvent ExecuteBegin(lpMessage)
    End Sub

    ''' <summary>
    ''' Wrapper for ExecuteComplete Event
    ''' </summary>
    ''' <param name="lpMessage"></param>
    ''' <remarks></remarks>
    Protected Sub RaiseExecuteComplete(ByVal lpMessage As String)
      RaiseEvent ExecuteComplete(lpMessage)
    End Sub

    ''' <summary>
    ''' Wrapper for ExecuteError Event
    ''' </summary>
    ''' <param name="lpMessage"></param>
    ''' <remarks></remarks>
    Protected Sub RaiseExecuteError(ByVal lpMessage As String)
      RaiseEvent ExecuteError(lpMessage)
    End Sub

    ''' <summary>
    ''' Wrapper for ExecuteReportProgress Event
    ''' </summary>
    ''' <param name="lpMessage"></param>
    ''' <remarks></remarks>
    Protected Sub RaiseExecuteReportProgress(ByVal lpPercentProgress As Integer, ByVal lpMessage As String)
      RaiseEvent ExecuteReportProgress(lpPercentProgress, lpMessage)
    End Sub

#End Region

#Region " IDisposable Support "

    Private disposedValue As Boolean     ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: free other state (managed objects).
        End If

        ' DISPOSETODO: free your own state (unmanaged objects).
        ' DISPOSETODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

#End Region

  End Class

End Namespace