' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperationException.vb
'  Description :  [type_description_here]
'  Created     :  12/1/2011 8:35:45 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

Public Class OperationException
  Inherits CtsException

#Region "Class Variables"

  Private mobjOperation As IOperable

#End Region

#Region "Public Properties"

  Public ReadOnly Property Operation As IOperable
    Get
      Return mobjOperation
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New(ByVal operation As IOperable, ByVal message As String)
    MyBase.New(message)
    Try
      mobjOperation = operation
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(ByVal operation As IOperable, ByVal message As String,
                 ByVal innerException As Exception)
    MyBase.New(message, innerException)
    Try
      mobjOperation = operation
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class
