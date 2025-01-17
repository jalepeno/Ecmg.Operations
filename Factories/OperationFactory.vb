' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperationFactory.vb
'  Description :  Used for creating operation objects
'  Created     :  11/29/2011 9:19:03 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Extensions
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Public Class OperationFactory
  Implements IDisposable
  'Implements ILoggable

#Region "Class Variables"

  Private Shared mintReferenceCount As Integer
  Private Shared mobjInstance As OperationFactory
  Private mobjAvailableOperations As IDictionary(Of String, Type) = Nothing
  Private mobjAvailableOperationDisplayNames As IList(Of String) = Nothing
  Private mobjAvailableSingleValuedParameterOperationDisplayNames As IList(Of String) = Nothing
  Private mobjAvailableCoreOperations As IDictionary(Of String, Type) = Nothing
  Private mobjAvailableExtensionOperations As IDictionary(Of String, Type) = Nothing
  Private mobjAvailableEnumParameters As IParameters = New Parameters
  Private mobjAvailableParameters As IParameters = New Parameters
  Private WithEvents mobjExtensionCatalog As ExtensionCatalog = ExtensionCatalog.Instance

#End Region

#Region "Constructors"

  Private Sub New()
    mintReferenceCount = 0
  End Sub

#End Region

#Region "Singleton Support"

  Public Shared Function Instance() As OperationFactory

    Try

      If mobjInstance Is Nothing Then
        mobjInstance = New OperationFactory
      End If

      mintReferenceCount += 1
      Return mobjInstance

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region

#Region "Public Properties"

  Public ReadOnly Property AvailableEnumParameters() As IParameters
    Get

      Try

        Return mobjAvailableEnumParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property AvailableParameters() As IParameters
    Get

      Try

        Return mobjAvailableParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property AvailableOperations() As IDictionary(Of String, Type)
    Get

      Try

        If mobjAvailableOperations Is Nothing Then
          mobjAvailableOperations = GetAllAvailableOperations()
        End If

        Return mobjAvailableOperations

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property AvailableOperationDisplayNames As IList(Of String)
    Get
      Try
        If mobjAvailableOperationDisplayNames Is Nothing Then
          mobjAvailableOperationDisplayNames = GetAllAvailableOperationDisplayNames()
        End If
        Return mobjAvailableOperationDisplayNames
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property AvailableSingleValuedParameterOperationDisplayNames As IList(Of String)
    Get
      Try
        If mobjAvailableSingleValuedParameterOperationDisplayNames Is Nothing Then
          mobjAvailableSingleValuedParameterOperationDisplayNames = GetAllAvailableOperationDisplayNames(True)
        End If
        Return mobjAvailableSingleValuedParameterOperationDisplayNames
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property AvailableCoreOperations As IDictionary(Of String, Type)
    Get
      Try
        If mobjAvailableCoreOperations Is Nothing Then
          mobjAvailableCoreOperations = GetAvailableCoreOperations()
        End If

        Return mobjAvailableCoreOperations
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property AvailableExtensionOperations As IDictionary(Of String, Type)
    Get
      Try
        If mobjAvailableExtensionOperations Is Nothing Then
          mobjAvailableExtensionOperations = GetAvailableExtensionOperations()
        End If

        Return mobjAvailableExtensionOperations
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  '  Protected Friend ReadOnly Property LogSession As Gurock.SmartInspect.Session Implements ILoggable.LogSession
  '  Get
  '    Try
  '      If mobjLogSession Is Nothing Then
  '        InitializeLogSession()
  '      End If
  '      Return mobjLogSession
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  'End Property

#End Region

#Region "Public Shared Methods"

  Public Shared Function Create(ByVal lpOperationType As String) As IOperable

    Try
      Return Instance.CreateOperation(lpOperationType, Result.NotProcessed, True, DateTime.MinValue, DateTime.MinValue)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Shared Function Create(ByVal lpOperationType As String,
                                ByVal lpResult As Result,
                                ByVal lpLogResult As Boolean,
                                ByVal lpStartTime As DateTime,
                                ByVal lpFinishTime As DateTime) As IOperable

    Try
      Return Instance.CreateOperation(lpOperationType, lpResult, lpLogResult, lpStartTime, lpFinishTime)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Friend Shared Function Create(lpOperationNode As XmlNode,
                                lpLocale As String) As IOperable

    Try
      Return Instance.CreateOperation(lpOperationNode, lpLocale)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Friend Shared Function Create(ByVal lpOperation As IOperable) As IOperable

    Try
      Return Instance.CreateOperation(lpOperation)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Shared Function GetAvailableOperationTypes() As IList(Of String)

    Try

      Dim lobjAvailableOperationTypes As New List(Of String)

      For Each lstrKey As String In Instance.GetAvailableCoreOperations.Keys
        lobjAvailableOperationTypes.Add(lstrKey)
      Next

      lobjAvailableOperationTypes.Sort()

      Return lobjAvailableOperationTypes

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Shared Function GetRunOperable(lpRunOperationNode As XmlNode) As IOperable

    Try

      If lpRunOperationNode.HasChildNodes = False Then
        Return Nothing

      Else

        If lpRunOperationNode.FirstChild.Name.ToLower.EndsWith("process") Then
          Return CType(Serializer.Deserialize.XmlString(lpRunOperationNode.FirstChild.OuterXml, GetType(Process)), IOperable)

        Else
          Return OperationFactory.Create(lpRunOperationNode.FirstChild, CultureInfo.CurrentCulture.Name)
        End If

      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region

#Region "Private Methods"

  Private Function CreateOperation(ByVal lpOperationType As String,
                                   ByVal lpResult As Result,
                                   ByVal lpLogResult As Boolean,
                                   ByVal lpStartTime As DateTime,
                                   ByVal lpFinishTime As DateTime) As IOperable

    Dim lobjOperationType As Type = Nothing
    Dim lobjTypes As Object = Nothing
    Dim lobjRequestedOperationType As Type = Nothing
    Dim lobjOperation As IOperation = Nothing
    Dim lstrOperationTypeName As String = lpOperationType.Replace(" ", String.Empty)

    Try

      If lpOperationType.ToLower = "migrate" Then
        Return ProcessFactory.CreateMigrationProcess(Nothing)
      ElseIf lpOperationType.ToLower = "migratefolder" Then
        Return ProcessFactory.CreateMigrateFolderProcess(Nothing)
      ElseIf lpOperationType.ToLower = "migratecustomobject" Then
        Return ProcessFactory.CreateMigrateCustomObjectProcess(Nothing)
      End If

      If AvailableOperations.ContainsKey(lstrOperationTypeName) = False Then
        Throw New UnknownOperationException(lpOperationType)

      Else

        Dim lobjCoreOperationTypeDictionary As IDictionary(Of String, Type) = Instance.AvailableOperations()

        If lobjCoreOperationTypeDictionary IsNot Nothing AndAlso lobjCoreOperationTypeDictionary.ContainsKey(lstrOperationTypeName) Then

          lobjOperationType = lobjCoreOperationTypeDictionary.Item(lstrOperationTypeName)
          lobjOperation = CType(Activator.CreateInstance(lobjOperationType), IOperation)

          With lobjOperation
            .SetResult(lpResult)
            .LogResult = lpLogResult
            .StartTime = lpStartTime
            .FinishTime = lpFinishTime
          End With

          Return lobjOperation

        End If

        Throw New ArgumentException(String.Format("Unable to create core operation of type '{0}': no operation defined with that name.", lpOperationType), "lpOperation")
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw

    Finally
      ' Clean up references
      lobjOperationType = Nothing
      lobjRequestedOperationType = Nothing
    End Try

  End Function

  Private Function CreateOperation(lpOperationNode As XmlNode,
                                   lpLocale As String) As IOperable

    Try

      Dim lobjOperation As IOperation
      Dim lstrOperationName As String = Nothing
      Dim lenuScope As OperationScope = OperationScope.Source
      Dim lenuResult As Result = Result.NotProcessed
      Dim lblnLogResult As Boolean = True
      Dim ldatStartTime As DateTime = DateTime.MinValue
      Dim ldatFinishTime As DateTime = DateTime.MinValue
      Dim lobjParameterNodes As XmlNodeList = Nothing
      Dim lobjFailureOperationNodes As XmlNodeList = Nothing
      Dim lobjFailureOperation As IOperation = Nothing
      Dim lobjParameter As IParameter = Nothing
      Dim lobjAttribute As XmlAttribute = Nothing
      Dim lobjNode As XmlNode = Nothing

      If lpOperationNode Is Nothing Then
        Throw New ArgumentNullException("lpOperationNode")
      End If

      If ((lpOperationNode.Name.StartsWith("Run")) AndAlso (Not lpOperationNode.Name.EndsWith("Operation"))) Then
        Return GetRunOperable(lpOperationNode)
      End If

      If lpOperationNode.HasChildNodes = False Then
        Throw New ArgumentException("Node has no child nodes", "lpOperationNode")
      End If

      ' Get the Name
      lobjAttribute = lpOperationNode.Attributes("Name")

      If lobjAttribute Is Nothing Then
        'Throw New ArgumentException("Node has no name attribute")
        lstrOperationName = lpOperationNode.Name

      Else
        lstrOperationName = lobjAttribute.InnerText
      End If

      ' Get the Scope
      lobjAttribute = lpOperationNode.Attributes("Scope")

      If lobjAttribute Is Nothing Then
        Throw New ArgumentException("Node has no Scope attribute")
      End If

      lenuScope = CType([Enum].Parse(GetType(OperationScope), lobjAttribute.InnerText), OperationScope)

      ' Get Result
      lobjAttribute = lpOperationNode.Attributes("Result")

      If lobjAttribute IsNot Nothing Then
        lenuResult = CType([Enum].Parse(GetType(Result), lobjAttribute.InnerText), Result)
      End If

      ' Get LogResult
      lobjAttribute = lpOperationNode.Attributes("LogResult")

      If lobjAttribute IsNot Nothing Then
        Boolean.TryParse(lobjAttribute.InnerText, lblnLogResult)
      End If

      ' Try to get the times

      ' Get the StartTime
      lobjAttribute = lpOperationNode.Attributes("StartTime")

      If lobjAttribute IsNot Nothing Then
        ldatStartTime = Helper.FromDetailedDateString(lobjAttribute.InnerText, lpLocale)
      End If

      ' Get the FinishTime
      lobjAttribute = lpOperationNode.Attributes("FinishTime")

      If lobjAttribute IsNot Nothing Then
        ldatFinishTime = Helper.FromDetailedDateString(lobjAttribute.InnerText, lpLocale)
      End If

      ' Create the operation
      lobjOperation = CType(OperationFactory.Create(lstrOperationName, lenuResult, lblnLogResult, ldatStartTime, ldatFinishTime), IOperation)
      lobjOperation.Scope = lenuScope

      If lobjOperation Is Nothing Then
        Throw New Exception(String.Format("Failed to create operation of type '{0}'", lstrOperationName))
      End If

      lobjParameterNodes = lpOperationNode.SelectNodes("Parameters/*")

      If lobjParameterNodes IsNot Nothing Then

        For Each lobjParameterNode As XmlNode In lobjParameterNodes
          lobjParameter = Nothing
          lobjParameter = Parameter.Create(lobjParameterNode)

          ' <Added by: Ernie at: 12/11/2014-1:35:32 PM on machine: ERNIE-THINK>
          ' This was added as a fix for some corrupted descriptions for the FolderDelimiter parameter of the ImportOperation.
          If TypeOf lobjOperation Is ImportOperation AndAlso lobjParameter.Name = "FolderDelimiter" Then
            lobjParameter.Description = ImportOperation.PARAM_FOLDER_DELIMITER_DESCRIPTION
          End If
          ' </Added by: Ernie at: 12/11/2014-1:35:32 PM on machine: ERNIE-THINK>

          If lobjParameter IsNot Nothing Then

            If lobjParameter.Type = PropertyType.ecmEnum Then
              ' Make sure we get the standard values if they are not already set
              If lobjParameter.HasStandardValues = False Then
                Dim lobjEnumParam As SingletonEnumParameter = lobjParameter

                Dim lobjOperationAssembly As Assembly
                Try
                  lobjOperationAssembly = Me.AvailableOperations(lstrOperationName).Assembly
                Catch ex As Exception
                  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
                  lobjOperationAssembly = Reflection.Assembly.GetExecutingAssembly
                End Try

                Dim lobjEnumType As System.Type = Helper.GetTypeFromAssembly(lobjOperationAssembly, lobjEnumParam.EnumName)

                If lobjEnumType IsNot Nothing Then
                  Dim lobjEnumDictionary As IDictionary(Of String, Integer) = Helper.EnumerationDictionary(lobjEnumType)
                  lobjEnumParam.SetStandardValues(lobjEnumDictionary.Keys)
                End If

              End If
            End If

            If lobjOperation.Parameters.Contains(lobjParameter.Name) Then
              lobjOperation.Parameters.Item(lobjParameter.Name) = lobjParameter

            Else
              lobjOperation.Parameters.Add(lobjParameter)
            End If

          End If

        Next

        ' Check the parameters in case there are any adjustments that are needed
        lobjOperation.CheckParameters()

      End If

      'lobjFailureOperationNodes = lpOperationNode.SelectNodes("FailureOperations/*")

      'If lobjFailureOperationNodes IsNot Nothing Then
      '  For Each lobjFailureOperationNode As XmlNode In lobjFailureOperationNodes
      '    lobjFailureOperation = Nothing
      '    lobjFailureOperation = CreateOperation(lobjFailureOperationNode, lpLocale)
      '    If lobjFailureOperation IsNot Nothing Then
      '      lobjOperation.OnFailureOperations.Add(lobjFailureOperation)
      '    End If
      '  Next
      'End If

      Dim lobjRunBeforeBeginNode As XmlNode = lpOperationNode.SelectSingleNode("RunBeforeBegin")

      If lobjRunBeforeBeginNode IsNot Nothing Then
        lobjOperation.RunBeforeBegin = Process.GetOperable(lobjRunBeforeBeginNode)
      End If

      Dim lobjRunAfterCompleteNode As XmlNode = lpOperationNode.SelectSingleNode("RunAfterComplete")

      If lobjRunAfterCompleteNode IsNot Nothing Then
        lobjOperation.RunAfterComplete = Process.GetOperable(lobjRunAfterCompleteNode)
      End If

      Dim lobjRunOnFailureNode As XmlNode = lpOperationNode.SelectSingleNode("RunOnFailure")

      If lobjRunOnFailureNode IsNot Nothing Then
        lobjOperation.RunOnFailure = Process.GetOperable(lobjRunOnFailureNode)
      End If

      If TypeOf lobjOperation Is IDecisionOperation Then

        ' TODO: Read the 'TrueOperations' node and the 'FalseOperations' node and update the respective collections.
        Dim lobjTrueOperationsNode As XmlNode = lpOperationNode.SelectSingleNode("TrueOperations")

        If lobjTrueOperationsNode IsNot Nothing Then
          CType(lobjOperation, IDecisionOperation).TrueOperations.AddRange(Process.GetOperations(lobjTrueOperationsNode))
        End If

        Dim lobjFalseOperationsNode As XmlNode = lpOperationNode.SelectSingleNode("FalseOperations")

        If lobjFalseOperationsNode IsNot Nothing Then
          CType(lobjOperation, IDecisionOperation).FalseOperations.AddRange(Process.GetOperations(lobjFalseOperationsNode))
        End If

      End If

      Return lobjOperation

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Private Function CreateOperation(ByVal lpOperation As IOperable) As IOperable

    Try

      ' This is used to clone an operation
      Dim lobjOperation As IOperation

      ' Create the operation
      lobjOperation = CType(OperationFactory.Create(lpOperation.Name, lpOperation.Result, lpOperation.LogResult, lpOperation.StartTime, lpOperation.FinishTime), IOperation)

      ' Add the parameters
      lobjOperation.Parameters = CType(lpOperation.Parameters.Clone, IParameters)

      If TypeOf lpOperation Is IDecisionOperation Then
        ' Clone the child operations as well
        CType(lobjOperation, IDecisionOperation).TrueOperations = CType(CType(lpOperation, IDecisionOperation).TrueOperations.Clone, IOperations)
        CType(lobjOperation, IDecisionOperation).FalseOperations = CType(CType(lpOperation, IDecisionOperation).FalseOperations.Clone, IOperations)
      End If

      Return lobjOperation

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Private Function GetAllAvailableOperationDisplayNames(Optional lpExcludeOperationsWithMultiValuedParameters As Boolean = False) As IList(Of String)
    Try
      Dim lobjReturnList As New List(Of String)

      If Not lpExcludeOperationsWithMultiValuedParameters Then
        For Each lstrKey As String In AvailableOperations.Keys
          lobjReturnList.Add(Helper.CreateDisplayName(lstrKey))
        Next
      Else
        Dim lobjOperation As IOperation = Nothing
        Dim lobjOperationType As Type = Nothing
        For Each lstrKey As String In AvailableOperations.Keys
          lobjOperationType = AvailableOperations(lstrKey)
          lobjOperation = CType(Activator.CreateInstance(lobjOperationType), IOperation)

          If lobjOperation IsNot Nothing Then
            ' Added by Ernie Bahr on 11/20/2015
            ' Until we can better handle these in the UI components we will filter out any operations with multi-valued parameters
            If Not ContainsMultiValuedParameter(lobjOperation) Then
              lobjReturnList.Add(Helper.CreateDisplayName(lstrKey))
            Else
              'LogSession.LogWarning("Operation '{0}' skipped over.  Operations with multi-valued parameters are currently unsupported.", lobjOperation.Name)
            End If
          End If

          lobjOperation = Nothing

        Next

      End If


      '' <Added by: Ernie at: 9/9/2014-4:01:49 PM on machine: ERNIE-THINK>
      '' Sumi wanted to be able to filter operations for the Box demo
      '' This is a quick and dirty implementation, we will improve or remove it later.
      '' Remove all excluded operations
      'Dim lstrProposedOperationKey As String
      'For lintOperationCounter As Integer = lobjReturnList.Count - 1 To 0 Step -1
      '  lstrProposedOperationKey = lobjReturnList(lintOperationCounter)
      '  If ConnectionSettings.Instance.OperationExclusions.Contains(lstrProposedOperationKey) Then
      '    lobjReturnList.Remove(lstrProposedOperationKey)
      '  End If
      'Next
      '' </Added by: Ernie at: 9/9/2014-4:01:49 PM on machine: ERNIE-THINK>

      Return lobjReturnList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetAllAvailableOperations() As IDictionary(Of String, Type)

    Try

      Dim lobjAllOperationList As New SortedDictionary(Of String, Type)
      Dim lobjCoreOperationList As IDictionary(Of String, Type) = AvailableCoreOperations
      Dim lobjExtensionOperationList As IDictionary(Of String, Type) = AvailableExtensionOperations

      For Each lstrCoreOperation As String In lobjCoreOperationList.Keys

        If lobjAllOperationList.ContainsKey(lstrCoreOperation) = False Then
          lobjAllOperationList.Add(lstrCoreOperation, lobjCoreOperationList.Item(lstrCoreOperation))
        End If

      Next

      For Each lstrExtensionOperation As String In lobjExtensionOperationList.Keys

        If lobjAllOperationList.ContainsKey(lstrExtensionOperation) = False Then
          lobjAllOperationList.Add(lstrExtensionOperation, lobjExtensionOperationList.Item(lstrExtensionOperation))
        End If

      Next

      Return lobjAllOperationList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Private Function GetAvailableCoreOperations() As IDictionary(Of String, Type)

    Dim lobjOperationList As New SortedDictionary(Of String, Type)
    Dim lobjOperationTypes As List(Of Type)
    Dim lobjOperation As IOperationInformation = Nothing

    Try

      lobjOperationTypes = GetAllCoreOperationTypes()

      For Each lobjType As Type In lobjOperationTypes

        lobjOperation = CType(Activator.CreateInstance(lobjType), IOperationInformation)

        If lobjOperation IsNot Nothing Then
          lobjOperationList.Add(lobjOperation.Name, lobjType)
          mobjAvailableEnumParameters.AddRange(GetAllEnumParameters(lobjOperation))
          mobjAvailableParameters.AddRange(GetAllParameters(lobjOperation))
        End If

        lobjOperation = Nothing

      Next

      Return lobjOperationList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw

    Finally
      ' Clean up references
      lobjOperationTypes = Nothing
    End Try

  End Function

  Private Function GetAllCoreOperationTypes() As List(Of Type)

    Dim lobjOperationList As New List(Of String)
    Dim lobjOperationType As Type = Nothing
    Dim lobjAssembly As Reflection.Assembly = Nothing
    Dim lobjTypes As IEnumerable(Of Type) = Nothing
    Dim lobjCoreOperationTypes As New List(Of Type)
    Dim lobjOperation As IOperation = Nothing

    Try

      'LogSession.EnterMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))

      lobjOperationType = GetType(IOperation)
      lobjAssembly = Reflection.Assembly.GetAssembly(lobjOperationType)
      lobjTypes = lobjAssembly.GetTypes.Where(Function(t) lobjOperationType.IsAssignableFrom(t))
      lobjOperation = Nothing

      For Each lobjType As Type In lobjTypes

        If lobjType.IsAbstract Then
          Continue For
        End If

        If lobjType.IsInterface Then
          Continue For
        End If

        lobjOperation = CType(Activator.CreateInstance(lobjType), IOperation)

        'If lobjOperation IsNot Nothing Then
        '  ' Added by Ernie Bahr on 11/20/2015
        '  ' Until we can better handle these in the UI components we will filter out any operations with multi-valued parameters
        '  If Not ContainsMultiValuedParameter(lobjOperation) Then
        '    lobjCoreOperationTypes.Add(lobjType)
        '  Else          
        '    'LogSession.LogWarning("Operation '{0}' skipped over.  Operations with multi-valued parameters are currently unsupported.", lobjOperation.Name)
        '  End If         
        'End If

        If lobjOperation IsNot Nothing Then
          lobjCoreOperationTypes.Add(lobjType)
        End If

        lobjOperation = Nothing

      Next

      Return lobjCoreOperationTypes

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw

    Finally
      ' Clean up references
      lobjAssembly = Nothing
      lobjTypes = Nothing
      lobjOperationType = Nothing
      'LogSession.LeaveMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
    End Try

  End Function

  'Private Function GetAllEnumParameters() As Parameters
  '  Try
  '    Dim lobjParameters As New Parameters
  '    For Each lobjOperation In AvailableOperations
  '      lobjParameters.AddRange(GetAllEnumParameters(lobjOperation))
  '    Next
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  Private Shared Function GetAllParameters(lpOperation As IOperation) As Parameters
    Try
      Dim lobjParameters As New Parameters

      For Each lobjParameter As IParameter In lpOperation.Parameters
        lobjParameters.Add(lobjParameter)
      Next

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Shared Function GetAllEnumParameters(lpOperation As IOperation) As Parameters
    Try
      Dim lobjParameters As New Parameters

      For Each lobjParameter As IParameter In lpOperation.Parameters
        If lobjParameter.Type = PropertyType.ecmEnum Then
          lobjParameters.Add(lobjParameter)
        End If
      Next

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Shared Function ContainsMultiValuedParameter(lpOperation As IOperation) As Boolean
    Try
      Return lpOperation.Parameters.Any(Function(lobjParameter) lobjParameter.Cardinality = Cardinality.ecmMultiValued)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>Finds all of the valid operation extension dlls in the specified folder path.</summary>
  ''' <param name="lpFolderPath">The folder path in which to search for operation extensions.</param>
  ''' <returns>A list of valid operation extensions.</returns>
  ''' <exception caption="FolderDoesNotExistException" cref="Exceptions.FolderDoesNotExistException">
  ''' If the specified folder does not exist, a FolderDoesNotExistException will be thrown.</exception>
  Public Shared Function GetAvailableExtensions(lpFolderPath As String) As IList(Of String)

    Try

      If Directory.Exists(lpFolderPath) = False Then
        Throw New FolderDoesNotExistException(lpFolderPath,
          String.Format("Unable to get available extensions, the path '{0}' is not valid.", lpFolderPath))
      End If

      Dim lobjExtensionList As New List(Of String)
      Dim lobjAssembly As System.Reflection.Assembly
      Dim lobjExtensionCandidate As Type

      ' Loop through each dll in the folder and see if it is an operation extension.
      For Each lstrExtensionCandidate As String In Directory.GetFiles(lpFolderPath, "*.dll")
        lobjExtensionCandidate = Nothing
        Try
          lobjAssembly = System.Reflection.Assembly.LoadFrom(lstrExtensionCandidate)
          'lobjExtensionCandidate = lobjAssembly. .GetType("OperationExtension")
          For Each lobjType As Type In lobjAssembly.GetExportedTypes
            If lobjType.IsAbstract = False Then
              lobjExtensionCandidate = lobjType.GetInterface("IOperationExtension")
              If lobjExtensionCandidate IsNot Nothing Then
                If lobjExtensionList.Contains(lstrExtensionCandidate) = False Then
                  lobjExtensionList.Add(lstrExtensionCandidate)
                End If
                Continue For
              End If
            End If
          Next
        Catch BadImageEx As BadImageFormatException
          ApplicationLogging.LogException(BadImageEx, Reflection.MethodBase.GetCurrentMethod)
          Continue For
        Catch ReflexLoadEx As ReflectionTypeLoadException
          ApplicationLogging.LogException(ReflexLoadEx, Reflection.MethodBase.GetCurrentMethod)
          Continue For
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Continue For
        End Try
      Next

      Return lobjExtensionList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Shared Function GetAvailableExtensionOperations(lpFolderPath As String) As IList(Of IExtensionInformation)

    Dim lobjExtensionList As IList(Of String)
    Dim lobjExtensionOperationList As New List(Of IExtensionInformation)
    'Dim lobjExtensionOperations As IEnumerable(Of Type)
    Dim lobjExportedTypes As Type() = Nothing
    Dim lobjExportedInterfaces As Type() = Nothing
    Dim lobjOperationExtensionTypes As Type() = Nothing
    Dim lobjOperationExtensionType As Type = Nothing
    Dim lobjExtensionAssembly As Assembly = Nothing
    Dim lobjExtensionInterface As Type = Nothing
    Dim lobjExtensionInstance As Extensions.IOperationExtension = Nothing

    Try

      lobjExtensionList = GetAvailableExtensions(lpFolderPath)
      'lobjOperationExtensionType = GetType(OperationExtension)
      For Each lstrExtension As String In lobjExtensionList
        lobjExtensionAssembly = Assembly.LoadFrom(lstrExtension)
        lobjExportedTypes = lobjExtensionAssembly.GetExportedTypes()
        For Each lobjExportedType As Type In lobjExportedTypes
          If lobjExportedType.IsAbstract Then
            Continue For
          End If
          lobjExtensionInterface = Nothing
          For Each lobjExportedInterface As Type In lobjExportedType.GetInterfaces
            If lobjExportedInterface.Name = "IOperationExtension" Then
              lobjExtensionInstance = CType(lobjExtensionAssembly.CreateInstance(lobjExportedType.Name, True), Extensions.IOperationExtension)
              'If lobjExtensionInstance IsNot Nothing Then
              '  lobjExtensionOperationList.Add( _
              '    New ExtensionInformation(lobjExportedType.Name.Replace("Operation", String.Empty), _
              '                             String.Empty, lobjExtensionInstance.CompanyName, _
              '                             lobjExtensionInstance.ProductName, lstrExtension))
              'Else
              '  lobjExtensionOperationList.Add( _
              '    New ExtensionInformation(lobjExportedType.Name.Replace("Operation", String.Empty), _
              '                             String.Empty, String.Empty, String.Empty, lstrExtension))
              'End If

              If lobjExtensionInstance IsNot Nothing Then
                ' Added by Ernie Bahr on 11/20/2015
                ' Until we can better handle these in the UI components we will filter out any operations with multi-valued parameters
                If Not ContainsMultiValuedParameter(lobjExtensionInstance) Then
                  lobjExtensionOperationList.Add(New ExtensionInformation(lobjExportedType.Name.Replace("Operation", String.Empty),
                                           String.Empty, lobjExtensionInstance.CompanyName,
                                           lobjExtensionInstance.ProductName, lstrExtension))
                Else
                  ' 'LogSession.LogWarning("Extension operation '{0}' skipped over.  Extension operations with multi-valued parameters are currently unsupported.", lobjExtensionInstance.Name)
                End If
              Else
                lobjExtensionOperationList.Add(
                  New ExtensionInformation(lobjExportedType.Name.Replace("Operation", String.Empty),
                                           String.Empty, String.Empty, String.Empty, lstrExtension))
              End If

            End If
          Next
        Next
      Next

      Return lobjExtensionOperationList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetAvailableExtensionOperations() As IDictionary(Of String, Type)

    Dim lobjExtensionOperationList As New SortedDictionary(Of String, Type)
    Dim lobjExtensionOperations As IEnumerable(Of KeyValuePair(Of String, Type))
    Dim lobjOperationType As Type = Nothing
    Dim lobjOperation As IOperationInformation = Nothing

    Try

      lobjOperationType = GetType(IOperation)
      lobjExtensionOperations = ExtensionCatalog.Instance.AvailableExtensions.Where(Function(t) lobjOperationType.IsAssignableFrom(t.Value))

      For Each lobjExtensionPair As KeyValuePair(Of String, Type) In lobjExtensionOperations

        lobjOperation = CType(Activator.CreateInstance(lobjExtensionPair.Value), IOperationInformation)

        If lobjOperation IsNot Nothing Then
          lobjExtensionOperationList.Add(lobjExtensionPair.Key, lobjExtensionPair.Value)
          mobjAvailableEnumParameters.AddRange(GetAllEnumParameters(lobjOperation))
          mobjAvailableParameters.AddRange(GetAllParameters(lobjOperation))
        End If

        lobjOperation = Nothing

      Next

      Return lobjExtensionOperationList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Private Sub RefreshCollections()
    Try
      mobjAvailableCoreOperations = GetAvailableCoreOperations()
      mobjAvailableExtensionOperations = GetAvailableExtensionOperations()
      mobjAvailableOperations = GetAllAvailableOperations()
    Catch Ex As Exception
      ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

  Private Sub mobjExtensionCatalog_CollectionChanged(sender As Object, e As Specialized.NotifyCollectionChangedEventArgs) Handles mobjExtensionCatalog.CollectionChanged
    Try
      RefreshCollections()
    Catch Ex As Exception
      ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  'Public Sub InitializeLogSession() Implements ILoggable.InitializeLogSession
  '  Try
  '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, System.Drawing.Color.PowderBlue)
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub

  'Public Sub FinalizeLogSession() Implements ILoggable.FinalizeLogSession
  '  Try
  '    If mobjLogSession IsNot Nothing Then
  '      ApplicationLogging.FinalizeLogSession(mobjLogSession)
  '    End If
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
        'FinalizeLogSession()
      End If

      ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' TODO: set large fields to null.
    End If
    disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    ' TODO: uncomment the following line if Finalize() is overridden above.
    ' GC.SuppressFinalize(Me)
  End Sub
#End Region
End Class
