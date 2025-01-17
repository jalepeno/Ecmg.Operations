'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Providers

Namespace PlugIns
  Public Class PlugInPropertyDefinition
    Inherits ProviderProperty

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpPropertyName As String,
                   ByVal lpPropertyType As Type,
                   Optional ByVal lpRequired As Boolean = True,
                   Optional ByVal lpPropertyValue As String = "",
                   Optional ByVal lpSequenceNumber As Integer = -1,
                    Optional ByVal lpDescription As String = "")
      MyBase.New(lpPropertyName, lpPropertyType, lpRequired, lpPropertyValue, lpSequenceNumber, lpDescription)
    End Sub
  End Class
End Namespace

