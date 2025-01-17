' ********************************************************************************
' '  Document    :  ItemProcessedEventArgs.vb
' '  Description :  [type_description_here]
' '  Created     :  10/2/2012-11:31:49
' '  <copyright company="ECMG">
' '      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

Imports Documents.Utilities

#End Region

Public Class ItemProcessedEventArgs

#Region "Class Variables"

  Private mobjWorkItem As IWorkItem = Nothing

#End Region

#Region "Public Properties"

  Public ReadOnly Property WorkItem As IWorkItem
    Get
      Try
        Return mobjWorkItem
      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New(lpWorkItem As IWorkItem)
    Try
      mobjWorkItem = lpWorkItem
    Catch Ex As Exception
      ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class
