' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IOperableResults.vb
'  Description :  [type_description_here]
'  Created     :  12/8/2011 8:02:22 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Public Interface IOperableResults
  Inherits ICollection(Of IOperableResult)
  Inherits IDisposable
  ' Property Item(name As String) As IOperableResult
  Function GetItemByName(name As String) As IOperableResult
  Sub SetItemByName(name As String, value As IOperableResult)
  Function ToXmlElementString() As String

End Interface
