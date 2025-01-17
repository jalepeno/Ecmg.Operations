' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IOperableResultsSummary.vb
'  Description :  [type_description_here]
'  Created     :  4/26/2016 10:49:40 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Public Interface IOperableResultsSummary
  Inherits ICollection(Of IOperableResultSummary)
  Inherits IDisposable

  Function GetItemByName(name As String) As IOperableResultSummary
  Sub SetItemByName(name As String, value As IOperableResultSummary)
  Function ToXmlElementString() As String

End Interface
