' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IProcessResult.vb
'  Description :  [type_description_here]
'  Created     :  12/8/2011 8:01:40 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Public Interface IProcessResult

  Inherits IOperableResult

  Property Node As String

  Property ContentCount As Integer

  Property TotalContentSize As FileSize

  Property VersionCount As Integer

  Property OperationResults As IOperableResults

  Property RunBeforeJobBeginResults As IOperableResults
  Property RunAfterJobCompleteResults As IOperableResults
  Property RunOnJobFailureResults As IOperableResults

End Interface
