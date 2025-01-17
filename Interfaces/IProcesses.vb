' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IProcesses.vb
'  Description :  [type_description_here]
'  Created     :  5/2/2012 10:51:05 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Public Interface IProcesses
  Inherits ICollection(Of IProcess)
  Inherits ICloneable

  Sub AddRange(ByVal lpProcesses As IEnumerable(Of IProcess))

End Interface
