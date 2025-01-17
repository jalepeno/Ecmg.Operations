'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Profiles

  <Serializable()>
  Public Class PlugInConfigurations
    Inherits CCollection(Of PlugIns.PlugInConfiguration)

#Region "Constructors"
    Public Sub New()

    End Sub
#End Region

  End Class

End Namespace