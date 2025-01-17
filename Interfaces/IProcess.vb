' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IProcess.vb
'  Description :  [type_description_here]
'  Created     :  11/23/2011 4:53:17 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Public Interface IProcess
  Inherits IOperable

  'ReadOnly Property DisplayName As String
  ReadOnly Property Operations As IOperations
  ReadOnly Property IsEmpty As Boolean
  Overloads ReadOnly Property ResultDetail As IProcessResult
  Function ToXmlString() As String

  Property RunBeforeParentBegin As IOperable
  Property RunAfterParentComplete As IOperable
  Property RunOnParentFailure As IOperable

  Property RunBeforeJobBegin As IOperable
  Property RunAfterJobComplete As IOperable
  Property RunOnJobFailure As IOperable

  ' Rollback Support
  Property AutoRollback As Boolean

End Interface