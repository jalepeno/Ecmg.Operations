' ********************************************************************************
' '  Document    :  IProcessResults.vb
' '  Description :  [type_description_here]
' '  Created     :  11/13/2012-16:14:30
' '  <copyright company="ECMG">
' '      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

#End Region

Public Interface IProcessResults
  Inherits ICollection(Of IProcessResult)

  'Property Item(name As String) As IProcessResult
  Function GetItemByName(name As String) As IProcessResult
  Sub SetItemByName(name As String, value As IProcessResult)

  Function GetOperationResultsByName(name As String) As IOperableResults
  Function GetOperationResultsByIndex(index As Integer) As IOperableResults

  Function ToJsonString() As String
  Function ToXmlString() As String
  Function ToXmlElementString() As String

End Interface
