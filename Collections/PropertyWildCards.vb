' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  PropertyWildCards.vb
'  Description :  [type_description_here]
'  Created     :  6/20/2023 4:12:07 PM
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

Public Class PropertyWildCards
  Inherits CCollection(Of PropertyWildCard)

#Region "Constructors"

  Public Sub New()

  End Sub

  Public Sub New(lpDocument As Document)
    Try

      For Each lobjProperty As ECMProperty In lpDocument.Properties
        Dim lobjPropertyWildCard As New PropertyWildCard(lobjProperty, PropertyScope.DocumentProperty)
        Add(lobjPropertyWildCard)
      Next

      For Each lobjProperty As ECMProperty In lpDocument.LatestVersion.Properties
        Dim lobjPropertyWildCard As New PropertyWildCard(lobjProperty, PropertyScope.VersionProperty)
        Add(lobjPropertyWildCard)
      Next

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Methods"

  Public Function SubstitutePropertyWildCards(lpDestinationFolderPath As String) As String
    Try

      Dim lstrReturnPath As String = lpDestinationFolderPath

      For Each lobjWildCard As PropertyWildCard In Items
        If lstrReturnPath.Contains(lobjWildCard.WildCard) Then
          lstrReturnPath = lstrReturnPath.Replace(lobjWildCard.WildCard, lobjWildCard.Property.Value)
        End If
        If Not lstrReturnPath.Contains("%") Then
          Exit For
        End If
      Next

      Return lstrReturnPath

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class
