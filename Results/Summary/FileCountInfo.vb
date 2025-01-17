' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  FileCountInfo.vb
'  Description :  [type_description_here]
'  Created     :  4/30/2016 11:50:40 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"


Imports System.Reflection
Imports Documents.Utilities

#End Region

Public Class FileCountInfo
  Inherits DoubleStatistical

#Region "Constructors"

  Public Sub New()

  End Sub

  ''' <summary>
  ''' Constructs a new FileSizeInfo instance using a collection of bytes.
  ''' </summary>
  ''' <param name="lpValues"></param>
  Public Sub New(lpValues As IEnumerable(Of Double))
    Try
      GetStatistics(lpValues)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class
