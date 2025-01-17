' ********************************************************************************
' '  Document    :  ProcessResultSummaries.vb
' '  Description :  [type_description_here]
' '  Created     :  9/12/2024-09:27:20
' '  <copyright company="Conteage">
' '      Copyright (c) Conteage Corp, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

Imports Documents.Core

#End Region

Public Class ProcessResultSummaries
  Inherits CCollection(Of IProcessResultSummary)
  Implements IProcessResultSummaries
  Implements IDisposable


  'Public Overloads Sub Add(item As IProcessResultSummary) Implements ICollection(Of IProcessResult).Add
  '  Try
  '    Add(CType(item, ProcessResult))
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub


#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        ' DISPOSETODO: dispose managed state (managed objects).
        MyBase.Dispose()
      End If

      ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' DISPOSETODO: set large fields to null.
    End If
    Me.disposedValue = True
  End Sub

#End Region

End Class
