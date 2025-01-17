' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  UnknownOperationException.vb
'  Description :  [type_description_here]
'  Created     :  2/21/2012 4:15:27 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Exceptions

#End Region

Public Class UnknownOperationException
  Inherits UnknownItemException

#Region "Public Properties"

  Public ReadOnly Property RequestedOperation As String
    Get
      Return MyBase.RequestedItem
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New(lpRequestedOperation As String)
    Me.New(String.Format(
           "Operation '{0}' unknown, if this is an extension operation make sure the extension is registered in ExtensionCatalog.xml.",
           lpRequestedOperation), lpRequestedOperation)
  End Sub

  Public Sub New(lpMessage As String, lpRequestedOperation As String)
    MyBase.New(lpMessage, lpRequestedOperation)
  End Sub

#End Region

End Class
