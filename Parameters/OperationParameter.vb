' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperationParameter.vb
'  Description :  Used to provide specific operation instructions.
'  Created     :  11/17/2011 4:32:15 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

Public Class OperationParameter
  Inherits Parameter
  Implements IOperationParameter

#Region "Constructors"

  Public Sub New()
    MyBase.New()
  End Sub

  Public Sub New(ByVal lpMenuType As PropertyType,
               ByVal lpName As String,
               ByVal lpValue As Object)
    MyBase.New(lpMenuType, lpName, lpValue)
  End Sub

#End Region

End Class