' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IOperations.vb
'  Description :  [type_description_here]
'  Created     :  11/23/2011 4:36:02 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Public Interface IOperations
  Inherits ICollection(Of IOperable)
  Inherits ICloneable
  'Inherits IOperable

  Sub AddRange(ByVal lpOperations As IEnumerable(Of IOperable))
  Sub AssignHost(lpHost As Object)
  Sub AssignTag(lpTag As Object)
  Function Execute(ByVal lpWorkItem As IWorkItem) As OperationEnumerations.Result
  Function Rollback(ByVal lpWorkItem As IWorkItem) As OperationEnumerations.Result
  Function ContainsOperation(lpOperationName As String) As Boolean
  Function ContainsOperation(lpOperationName As String, ByRef lpFoundOperation As IOperable) As Boolean

End Interface