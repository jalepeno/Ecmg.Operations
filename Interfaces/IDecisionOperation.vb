' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IDecisionOperation.vb
'  Description :  [type_description_here]
'  Created     :  4/13/2012 10:55:24 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Public Interface IDecisionOperation
  Inherits IOperation

  Function Evaluate() As Boolean
  ReadOnly Property Evaluation As Boolean
  Property TrueOperations As IOperations
  Property FalseOperations As IOperations
  ReadOnly Property RunOperations As IOperations

End Interface