' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IOperation.vb
'  Description :  [type_description_here]
'  Created     :  11/18/2011 3:13:42 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Public Interface IOperation
  Inherits IOperable

  ''' <summary>
  ''' Used to indicate the scope of the operation.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Property Scope As OperationScope
  Property ScopeString As String

End Interface