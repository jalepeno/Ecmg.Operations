'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ProcessEventArgs.vb
'   Description :  [type_description_here]
'   Created     :  6/27/2013 11:16:41 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Arguments
Imports Documents.Utilities

#End Region

Public Class ProcessEventArgs
  Inherits ItemEventArgs

#Region "Class Variables"

  Private mobjProcess As IProcess = Nothing

#End Region

#Region "Public Properties"

  Public ReadOnly Property Process As IProcess
    Get
      Return mobjProcess
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New()
    MyBase.New(Nothing)
  End Sub

  Public Sub New(lpProcess As IProcess)
    MyBase.New(lpProcess)
    Try
      mobjProcess = lpProcess
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class
