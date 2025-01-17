'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core
Imports Documents.Providers

Namespace PlugIns
  Public Class PlugInPropertyDefinitions
    Inherits ProviderProperties

    ''' <summary>
    ''' Converts a set of definitions into a set of PlugInProperties
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ToPlugInProperties() As PlugInProperties


      Try
        Dim lobjPlugInProperties As New PlugInProperties


        For Each lobjPlugInDef As PlugInPropertyDefinition In Me
          'Fill it in
          Dim lobjPlugInProperty As New PlugInProperty With {
            .Cardinality = Cardinality.ecmSingleValued,
            .Name = lobjPlugInDef.PropertyName,
            .Value = lobjPlugInDef.PropertyValue,
            .Type = SystemTypeToEcmgType(lobjPlugInDef.PropertyType),
            .Description = lobjPlugInDef.Description
          }

          lobjPlugInProperties.Add(lobjPlugInProperty)

        Next

        Return lobjPlugInProperties

      Catch ex As Exception
        Throw
      End Try

      Return Nothing

    End Function

    Private Shared Function SystemTypeToEcmgType(ByVal lpSystemType As System.Type) As PropertyType

      Select Case lpSystemType.ToString

        Case "System.String"
          Return PropertyType.ecmString

        Case "System.Boolean"
          Return PropertyType.ecmBoolean

        Case "System.Object"
          Return PropertyType.ecmObject

        Case "System.DateTime"
          Return PropertyType.ecmDate

        Case "System.Double"
          Return PropertyType.ecmDouble

        Case "System.Guid"
          Return PropertyType.ecmGuid

        Case Else
          Return PropertyType.ecmObject


      End Select

    End Function

  End Class
End Namespace

