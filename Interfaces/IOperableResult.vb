' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IOperableResult.vb
'  Description :  [type_description_here]
'  Created     :  12/8/2011 7:56:40 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Public Interface IOperableResult
  Inherits IDisposable

  Property Name As String
  Property Result As OperationEnumerations.Result
  Property ProcessedMessage As String
  Property Scope As OperationScope
  Property StartTime As DateTime
  Property FinishTime As DateTime
  Property TotalProcessingTime As TimeSpan

  ReadOnly Property Parent As IOperable

  Property ChildOperations As IOperableResults

  Property RunBeforeBeginResults As IOperableResults
  Property RunAfterCompleteResults As IOperableResults
  Property RunOnFailureResults As IOperableResults
  Property RunBeforeParentBeginResults As IOperableResults
  Property RunAfterParentCompleteResults As IOperableResults
  Property RunOnParentFailureResults As IOperableResults

  Function ToJsonString() As String
  Function ToXmlString() As String
  Function ToXmlElementString() As String

End Interface
