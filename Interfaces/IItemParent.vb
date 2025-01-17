' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IItemParent.vb
'  Description :  [type_description_here]
'  Created     :  12/2/2011 9:08:58 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Transformations

#End Region

Public Interface IItemParent

  Property Id As String

  Property Name As String

  Property SourceConnection As IRepositoryConnection

  Property DestinationConnection As IRepositoryConnection

  Property ExportPath As String

  Property Transformations As TransformationCollection

  Delegate Sub ItemProcessed(ByVal sender As Object, ByRef e As ItemProcessedEventArgs)

  Sub RefreshSourceConnection()

  Sub RefreshDestinationConnection()

End Interface
