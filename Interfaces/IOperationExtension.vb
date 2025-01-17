' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IOperationExtension.vb
'  Description :  Used for managing operation specific extension dlls for projects.
'  Created     :  11/16/2011 10:13:05 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Extensions

#End Region

Namespace Extensions

  Public Interface IOperationExtension
    Inherits IExtension
    Inherits IOperation
  End Interface

End Namespace