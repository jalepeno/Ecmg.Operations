' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperationNotInitializedException.vb
'  Description :  [type_description_here]
'  Created     :  12/1/2011 8:31:15 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Public Class OperationNotInitializedException
  Inherits OperationException

#Region "Constructors"

  Public Sub New(ByVal operation As IOperable, ByVal message As String)
    MyBase.New(operation, message)
  End Sub

  Public Sub New(ByVal operation As IOperable, ByVal message As String,
                 ByVal innerException As Exception)
    MyBase.New(operation, message, innerException)
  End Sub

#End Region

End Class
