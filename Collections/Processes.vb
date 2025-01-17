' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Processes.vb
'  Description :  [type_description_here]
'  Created     :  5/2/2012 10:53:54 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class Processes
  Inherits CCollection(Of IProcess)
  Implements IProcesses
  Implements ICloneable

#Region "IProcesses Implementation"

  Public Overrides Sub AddRange(ByVal lpProcesses As System.Collections.Generic.IEnumerable(Of IProcess)) Implements IProcesses.AddRange
    Try
      MyBase.AddRange(lpProcesses)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "ICloneable Implementation"

  Public Function Clone() As Object Implements System.ICloneable.Clone
    Try
      Dim lobjProcesses As New Processes

      For Each lobjProcess As IProcess In Me
        lobjProcesses.Add(CType(lobjProcess.Clone, IProcess))
      Next

      Return lobjProcesses

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
