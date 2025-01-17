' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ItemParent.vb
'  Description :  [type_description_here]
'  Created     :  12/5/2011 4:10:50 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Transformations
Imports Documents.Utilities

#End Region

Public Class ItemParent
  Implements IItemParent

#Region "Class Variables"

  Private mstrId As String = String.Empty
  Private mstrName As String = String.Empty
  Private mstrSourceConnectionString As String = String.Empty
  Private mstrDestinationConnectionString As String = String.Empty
  Private mobjSourceConnection As IRepositoryConnection = Nothing
  Private mobjDestinationConnection As IRepositoryConnection = Nothing
  Private mstrExportPath As String = String.Empty
  Private mobjTransformations As New TransformationCollection

#End Region

#Region "Constructors"

  Public Sub New()

  End Sub

  Public Sub New(lpSourceConnectionString As String)
    Try
      mstrSourceConnectionString = lpSourceConnectionString
      SourceConnection = New ContentSource(lpSourceConnectionString)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(lpSourceConnectionString As String, lpDestinationConnectionString As String)
    Try

      mstrSourceConnectionString = lpSourceConnectionString
      SourceConnection = New ContentSource(lpSourceConnectionString)

      If Not String.IsNullOrEmpty(lpDestinationConnectionString) Then
        mstrDestinationConnectionString = lpDestinationConnectionString
        DestinationConnection = New ContentSource(lpDestinationConnectionString)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(lpSourceConnection As IRepositoryConnection)
    Try
      SourceConnection = lpSourceConnection
      mstrSourceConnectionString = lpSourceConnection.ConnectionString
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(lpSourceConnection As IRepositoryConnection, lpDestinationConnection As IRepositoryConnection)
    Try
      SourceConnection = lpSourceConnection
      mstrSourceConnectionString = lpSourceConnection.ConnectionString

      DestinationConnection = lpDestinationConnection
      mstrDestinationConnectionString = lpDestinationConnection.ConnectionString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "IItemParent Implementation"

  Public Property Id As String Implements IItemParent.Id
    Get
      Return mstrId
    End Get
    Set(value As String)
      mstrId = value
    End Set
  End Property

  Public Property Name As String Implements IItemParent.Name
    Get
      Return mstrName
    End Get
    Set(value As String)
      mstrName = value
    End Set
  End Property

  Public Property DestinationConnection As IRepositoryConnection Implements IItemParent.DestinationConnection
    Get
      Return mobjDestinationConnection
    End Get
    Set(value As IRepositoryConnection)
      mobjDestinationConnection = value
    End Set
  End Property

  Public Property ExportPath As String Implements IItemParent.ExportPath
    Get
      Return mstrExportPath
    End Get
    Set(value As String)
      mstrExportPath = value
    End Set
  End Property

  Public Property SourceConnection As IRepositoryConnection Implements IItemParent.SourceConnection
    Get
      Return mobjSourceConnection
    End Get
    Set(value As IRepositoryConnection)
      mobjSourceConnection = value
    End Set
  End Property

  Public Property Transformations As TransformationCollection Implements IItemParent.Transformations
    Get
      Return mobjTransformations
    End Get
    Set(value As TransformationCollection)
      mobjTransformations = value
    End Set
  End Property

  Public Sub RefreshDestinationConnection() Implements IItemParent.RefreshDestinationConnection
    Try
      If Not String.IsNullOrEmpty(mstrDestinationConnectionString) Then
        DestinationConnection = New ContentSource(mstrDestinationConnectionString)
      Else
        Throw New InvalidConnectionStringException("Unable to refresh destination connection, the connection string is not initialized.")
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub RefreshSourceConnection() Implements IItemParent.RefreshSourceConnection
    Try
      If Not String.IsNullOrEmpty(mstrSourceConnectionString) Then
        SourceConnection = New ContentSource(mstrSourceConnectionString)
      Else
        Throw New InvalidConnectionStringException("Unable to refresh source connection, the connection string is not initialized.")
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class
