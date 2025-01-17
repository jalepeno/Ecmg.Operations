'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core
Imports Documents.SerializationUtilities

Namespace PlugIns

  <Serializable()>
  Public Class PlugInConfiguration
    Implements IDescription
    Implements IDisposable

#Region "Class Variables"
    Private mobjPlugInProperties As PlugInProperties
    Private mstrPlugInPath As String = String.Empty
    Private mstrDescription As String = String.Empty
    Private mstrName As String = String.Empty
    'Private mstrVersion As String = String.Empty
    'Private mstrPlugInClassName As String
    'Private menumPlugInExecutionMode As PlugInExecutionMode
    'Private menuPlugInExecuteTiming As PlugInExecuteTiming
    'Private menuPlugInType As PlugInType

#End Region

#Region "Public Properties"

    <Xml.Serialization.XmlAttribute("Name")>
    Public Property Name() As String Implements IDescription.Name, INamedItem.Name
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    <Xml.Serialization.XmlAttribute("Description")>
    Public Property Description() As String Implements IDescription.Description
      Get
        Return mstrDescription
      End Get
      Set(ByVal value As String)
        mstrDescription = value
      End Set
    End Property

    Public Property PlugInProperties() As PlugInProperties
      Get
        Return mobjPlugInProperties
      End Get
      Set(ByVal value As PlugInProperties)
        mobjPlugInProperties = value
      End Set
    End Property

    Public Property PlugInPath() As String
      Get
        Return mstrPlugInPath
      End Get
      Set(ByVal value As String)
        mstrPlugInPath = value
      End Set
    End Property

    'Public Property PlugInClassName() As String
    '  Get
    '    Return mstrPlugInClassName
    '  End Get
    '  Set(ByVal value As String)
    '    mstrPlugInClassName = value
    '  End Set
    'End Property

    'Public Property Version() As String
    '  Get
    '    Return mstrVersion
    '  End Get
    '  Set(ByVal value As String)
    '    mstrVersion = value
    '  End Set
    'End Property

    'Public Property PlugInExecuteTiming() As PlugInExecuteTiming
    '  Get
    '    Return menuPlugInExecuteTiming
    '  End Get
    '  Set(ByVal value As PlugInExecuteTiming)
    '    menuPlugInExecuteTiming = value
    '  End Set
    'End Property

    'Public Property PlugInExecutionMode() As PlugInExecutionMode
    '  Get
    '    Return menumPlugInExecutionMode
    '  End Get
    '  Set(ByVal value As PlugInExecutionMode)
    '    menumPlugInExecutionMode = value
    '  End Set
    'End Property

    'Public Property PlugInType() As PlugInType
    '  Get
    '    Return menuPlugInType
    '  End Get
    '  Set(ByVal value As PlugInType)
    '    menuPlugInType = value
    '  End Set
    'End Property



#End Region

#Region "Private Properties"

    Private ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

#End Region

#Region "Public Methods"

    Sub Serialize(ByVal lpFilePath As String)
      Try
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function Serialize() As System.Xml.XmlDocument
      Try
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Deserialize(ByVal lpFilePath As String) As Object
      Try
        Return Serializer.Deserialize.XmlFile(lpFilePath, GetType(PlugInConfiguration))
      Catch ex As Exception
        Return Nothing
      End Try
    End Function

    'Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object
    '  Try
    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If
    '    Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
    '  Catch ex As Exception
    '    'ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpFilePath, lpErrorMessage))
    '    lpErrorMessage = ex.Message
    '    Return Nothing
    '  End Try
    'End Function

    Public Function DeSerialize(ByVal lpXML As System.Xml.XmlDocument) As Object
      Try
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        'Helper.DumpException(ex)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

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
