'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IOperationExtensions.vb
'   Description :  [type_description_here]
'   Created     :  1/9/2015 11:38:36 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Operations.Extensions

#End Region

Public Interface IOperationExtensions
  Inherits ICollection(Of IOperationExtension)
End Interface
