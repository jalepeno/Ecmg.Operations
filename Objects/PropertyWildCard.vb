' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  PropertyWildCard.vb
'  Description :  [type_description_here]
'  Created     :  6/20/2023 3:57:07 PM
'   <copyright company="Conteage">
'       Copyright (c) Conteage Corp. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class PropertyWildCard

#Region "Class Variables"

  Private mstrWildCard As String = String.Empty
  Private mobjProperty As ECMProperty
  Private menuScope As PropertyScope

#End Region

#Region "Public Properties"

  Public ReadOnly Property WildCard As String
    Get
      Return mstrWildCard
    End Get
  End Property

  Public ReadOnly Property [Property] As ECMProperty
    Get
      Return mobjProperty
    End Get
  End Property

  Public ReadOnly Property Scope As PropertyScope
    Get
      Return menuScope
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New(lpProperty As ECMProperty, lpScope As PropertyScope)
    Try
      mobjProperty = lpProperty
      menuScope = lpScope
      mstrWildCard = GetWildCard(lpProperty)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Methods"

  Public Overrides Function ToString() As String
    Try
      If Not String.IsNullOrEmpty(mobjProperty.Value) Then
        Return String.Format("{0}: {1}", mstrWildCard, mobjProperty.Value)
      Else
        Return String.Format("{0}: No value - {1}", mstrWildCard, mobjProperty.SystemName)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Function GetWildCard(lpProperty As ECMProperty) As String
    Try
      Dim lstrWildCard As String = String.Format("{0}{1}{0}", "%", lpProperty.SystemName)
      Return lstrWildCard
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
