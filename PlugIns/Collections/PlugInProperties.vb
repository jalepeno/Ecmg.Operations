'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities

Namespace PlugIns
  <Serializable()>
  <Xml.Serialization.XmlInclude(GetType(PlugInProperty))>
  Public Class PlugInProperties
    Inherits ECMProperties

#Region "Private Properties"

#End Region

#Region "Public Methods"

    Public Overloads Function Clone() As Object

      Dim lobjProperties As New PlugInProperties

      Try
        For Each lobjProperty As PlugInProperty In Me
          lobjProperties.Add(lobjProperty.Clone)
        Next
        Return lobjProperties
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
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
        Dim lobjPlugInProperties As PlugInProperties = CType(Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType), PlugInProperties)
        Return lobjPlugInProperties
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

#End Region

  End Class
End Namespace
