' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Operation.vb
'  Description :  [type_description_here]
'  Created     :  11/18/2011 4:05:34 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Globalization
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

<TypeConverter(GetType(ExpandableObjectConverter)),
DebuggerDisplay("{DebuggerIdentifier(),nq}")>
Public MustInherit Class Operation
  Implements IOperation
  Implements IOperationInformation
  Implements IJsonSerializable(Of IOperation)
  Implements IXmlSerializable
  Implements ICloneable
  'Implements ILoggable

#Region "Class Variables"

  Private mstrDocumentId As String = String.Empty
  Private mstrProcessedMessage As String = String.Empty
  Private mobjFailureOperations As New Operations
  Protected menuResult As Result = Result.NotProcessed
  Private mblnLogResult As Boolean = True
  Private mobjResultDetail As IOperableResult = Nothing
  Private menuScope As OperationScope
  Private WithEvents mobjParameters As IParameters = New Parameters
  Protected mstrDescription As String = String.Empty
  Private mobjOperableParent As IOperable = Nothing
  Private mobjParent As IItemParent = Nothing
  Private mobjWorkItem As IWorkItem = Nothing
  Private mdatStartTime As DateTime = DateTime.MinValue
  Private mdatFinishTime As DateTime = DateTime.MinValue
  Private mobjSourceConnection As IRepositoryConnection = Nothing
  Private mobjDestinationConnection As IRepositoryConnection = Nothing
  Private mobjPrimaryConnection As IRepositoryConnection = Nothing
  Private mobjLocale As CultureInfo = CultureInfo.CurrentCulture
  Private mstrLocale As String = mobjLocale.Name
  Private mobjHost As Object = Nothing
  Private mobjTag As Object = Nothing
  Private mobjRunBeforeBegin As IOperable = Nothing
  Private mobjRunAfterComplete As IOperable = Nothing
  Private mobjRunOnFailure As IOperable = Nothing
  Private mstrInstanceId As String = String.Empty

  Private WithEvents mobjParameterCollection As ObservableCollection(Of IParameter) = Nothing

  Private Shared mstrParameterExpression As String = "(?<Prefix>.*){(?<ParamName>[a-zA-Z0-9]*):(?<ParamValue>[a-zA-Z0-9]*)}(?<Suffix>.*)"

  'Private mobjLogSession As Gurock.SmartInspect.Session

#End Region

#Region "Class Constants"

  Friend Const LOG_RESULT As String = "LogResult"
  Friend Const CATEGORY_BEHAVIOR As String = "Behavior"
  Friend Const CATEGORY_CONFIG As String = "Configuration"

#End Region

#Region "Public Events"

  Public Event Begin(ByVal sender As Object, ByVal e As OperableEventArgs) Implements IOperation.Begin

  Public Event Complete(ByVal sender As Object, ByVal e As OperableEventArgs) Implements IOperation.Complete

  Public Event OperatingError(ByVal sender As Object, ByVal e As OperableErrorEventArgs) Implements IOperation.OperatingError
  Public Event ParameterPropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements IOperable.ParameterPropertyChanged

#End Region

#Region "Constructors"

  Public Sub New()
    Try
      If Parameters.Count = 0 Then
        Parameters = GetDefaultParameters()
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Properties"

  ''' <summary>
  ''' The name of the operation.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public MustOverride ReadOnly Property Name As String Implements IOperable.Name, IOperationInformation.Name

  Public MustOverride ReadOnly Property CanRollback As Boolean Implements IOperable.CanRollback

  Public Overridable ReadOnly Property DisplayName As String Implements IOperable.DisplayName, IOperationInformation.DisplayName
    Get
      Try
        Return Helper.CreateDisplayName(Me.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property CompanyName As String Implements IOperationInformation.CompanyName
    Get
      Return ConstantValues.CompanyName
    End Get
  End Property

  Public ReadOnly Property ProductName As String Implements IOperationInformation.ProductName
    Get
      Return ConstantValues.ProductName
    End Get
  End Property

  Public Overridable ReadOnly Property Description As String Implements IOperable.Description, IOperationInformation.Description
    Get
      Return mstrDescription
    End Get
    'Set(value As String)
    '  mstrDescription = value
    'End Set
  End Property

  Public ReadOnly Property IsDisposed() As Boolean Implements IOperation.IsDisposed
    Get
      Return disposedValue
    End Get
  End Property

  Public ReadOnly Property IsExtension As Boolean Implements IOperationInformation.IsExtension
    Get
      Return False
    End Get
  End Property

  Public Property ExtensionPath As String Implements IOperationInformation.ExtensionPath
    Get
      Return String.Empty
    End Get
    Set(ByVal value As String)

    End Set
  End Property

  ''' <summary>
  ''' The Id of the document to operate on.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property DocumentId As String Implements IOperation.DocumentId
    Get
      Return mstrDocumentId
    End Get
    Set(ByVal value As String)
      mstrDocumentId = value
    End Set
  End Property

  Public Property OperableParent As IOperable Implements IOperable.OperableParent
    Get
      Try
        Return mobjOperableParent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IOperable)
      Try
        mobjOperableParent = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Parent As IItemParent Implements IOperation.Parent
    Get
      Return mobjParent
    End Get
    Set(ByVal value As IItemParent)
      mobjParent = value
    End Set
  End Property

  Public Property WorkItem As IWorkItem Implements IOperation.WorkItem
    Get
      Return mobjWorkItem
    End Get
    Set(ByVal value As IWorkItem)
      Try
        mobjWorkItem = value
        If Parent Is Nothing AndAlso value.Parent IsNot Nothing Then
          Parent = value.Parent
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set

  End Property

  Public Property ProcessedMessage As String Implements IOperation.ProcessedMessage
    Get
      Return mstrProcessedMessage
    End Get
    Set(ByVal value As String)
      Try
        mstrProcessedMessage = value
        If WorkItem IsNot Nothing Then
          If String.IsNullOrEmpty(WorkItem.ProcessedMessage) Then
            WorkItem.ProcessedMessage = value
          Else
            WorkItem.ProcessedMessage = $"{WorkItem.ProcessedMessage}, {value}"
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  ''' <summary>
  ''' Indicates the result of the operation.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property Result As OperationEnumerations.Result Implements IOperation.Result
    Get
      Return menuResult
    End Get
  End Property

  Public ReadOnly Property ResultDetail As IOperableResult Implements IOperation.ResultDetail
    Get
      Return mobjResultDetail
    End Get
  End Property

  ''' <summary>
  ''' Gets or sets a value indicating whether or 
  ''' not the result of the operation should be logged.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property LogResult As Boolean Implements IOperable.LogResult
    Get
      Return mblnLogResult
    End Get
    Set(ByVal value As Boolean)
      mblnLogResult = value
    End Set
  End Property

  ''' <summary>
  ''' Gets or sets the scope of the operation as the source or the destination document.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  <Category(CATEGORY_CONFIG), Description("Gets or sets the scope of the operation as the source or the destination document.")>
  Public Property Scope As OperationEnumerations.OperationScope Implements IOperation.Scope
    Get
      Return menuScope
    End Get
    Set(ByVal value As OperationEnumerations.OperationScope)
      menuScope = value
    End Set
  End Property

  <Description("Gets or sets the scope of the operation as the source or the destination document.")>
  Public Property ScopeString As String Implements IOperation.ScopeString
    Get
      Try
        Return Me.Scope.ToString
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As String)
      Try
        If [Enum].TryParse(value, Me.Scope) = False Then
          Throw New InvalidEnumArgumentException(String.Format("{0} is not a valid OperationScope value", value))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  ''' <summary>
  ''' Gets or sets the parameters for the operation.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  <Category(CATEGORY_CONFIG)>
  Public Property Parameters As IParameters Implements IOperable.Parameters
    Get
      Return mobjParameters
    End Get
    Set(ByVal value As IParameters)
      mobjParameters = value
    End Set
  End Property

  Public Property ShouldExecute As Boolean Implements IOperable.ShouldExecute

  ' ''' <summary>
  ' ''' Gets or sets the collection of operations to execute if this operation fails
  ' ''' </summary>
  ' ''' <value></value>
  ' ''' <returns></returns>
  ' ''' <remarks></remarks>
  'Public Property FailureOperations As IOperations Implements IOperable.OnFailureOperations
  '  Get
  '    Try
  '      Return mobjFailureOperations
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  '  Set(ByVal value As IOperations)
  '    Try
  '      mobjFailureOperations = value
  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Set
  'End Property

  ''' <summary>
  ''' Gets or sets the time when the operation is started.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property StartTime As DateTime Implements IOperation.StartTime
    Get
      Return mdatStartTime
    End Get
    Set(ByVal value As DateTime)
      mdatStartTime = value
    End Set
  End Property

  ''' <summary>
  ''' Gets or sets the time when the operation is finished.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property FinishTime As DateTime Implements IOperation.FinishTime
    Get
      Return mdatFinishTime
    End Get
    Set(ByVal value As DateTime)
      mdatFinishTime = value
    End Set
  End Property

  ''' <summary>
  ''' Gets the total processing time for the operation.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TotalProcessingTime As TimeSpan Implements IOperation.TotalProcessingTime
    Get
      Try
        Return FinishTime - StartTime
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property Locale As CultureInfo Implements IOperable.Locale
    Get
      Return mobjLocale
    End Get
  End Property

  <XmlIgnore()>
  Public Property Host As Object Implements IOperable.Host
    Get
      Try
        Return mobjHost
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Object)
      Try
        mobjHost = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  <XmlIgnore()>
  Public Property Tag As Object Implements IOperable.Tag
    Get
      Try
        Return mobjTag
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Object)
      Try
        mobjTag = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  <Category(CATEGORY_BEHAVIOR)>
  Public Property RunBeforeBegin As IOperable Implements IOperable.RunBeforeBegin
    Get
      Return mobjRunBeforeBegin
    End Get
    Set(value As IOperable)
      mobjRunBeforeBegin = value
    End Set
  End Property

  <Category(CATEGORY_BEHAVIOR)>
  Public Property RunAfterComplete As IOperable Implements IOperable.RunAfterComplete
    Get
      Return mobjRunAfterComplete
    End Get
    Set(value As IOperable)
      mobjRunAfterComplete = value
    End Set
  End Property

  <Category(CATEGORY_BEHAVIOR)>
  Public Property RunOnFailure As IOperable Implements IOperable.RunOnFailure
    Get
      Return mobjRunOnFailure
    End Get
    Set(value As IOperable)
      mobjRunOnFailure = value
    End Set
  End Property

  Public ReadOnly Property InstanceId As String Implements IOperable.InstanceId
    Get
      Return mstrInstanceId
    End Get
  End Property

  'Public MustOverride ReadOnly Property CanRollback As Boolean Implements IOperable.CanRollback

#End Region

#Region "Friend Properties"

  'Protected Friend ReadOnly Property LogSession As Gurock.SmartInspect.Session Implements ILoggable.LogSession
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

  Friend ReadOnly Property SourceConnection As IRepositoryConnection
    Get
      Try

        If mobjSourceConnection Is Nothing Then
          mobjSourceConnection = GetConnection(OperationScope.Source)
        End If

        Return mobjSourceConnection

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Friend ReadOnly Property DestinationConnection As IRepositoryConnection
    Get
      Try

        If mobjDestinationConnection Is Nothing Then
          mobjDestinationConnection = GetConnection(OperationScope.Destination)
        End If

        Return mobjDestinationConnection

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Friend ReadOnly Property PrimaryConnection As IRepositoryConnection
    Get
      Try

        If mobjPrimaryConnection Is Nothing Then
          mobjPrimaryConnection = GetPrimaryConnection()
        End If

        Return mobjPrimaryConnection

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

#End Region

#Region "Protected Methods"

  'Protected Overridable Sub InitializeLogSession() Implements ILoggable.InitializeLogSession
  '  Try
  '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, System.Drawing.Color.PaleGreen)
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub

  'Protected Overridable Sub FinalizeLogSession() Implements ILoggable.FinalizeLogSession
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

  Protected Friend Overridable Function DebuggerIdentifier() As String
    Dim lobjIdentifierBuilder As New Text.StringBuilder
    Try

      If disposedValue = True Then
        If String.IsNullOrEmpty(Me.Name) Then
          Return "Operation Disposed"
        Else
          Return String.Format("{0} Operation Disposed", Me.Name)
        End If
      End If

      Dim lstrName As String = Name

      lobjIdentifierBuilder.AppendFormat("{0} Operation", Name)

      Return lobjIdentifierBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Return lobjIdentifierBuilder.ToString
    End Try
  End Function

  'Protected MustOverride Function OnRollback() As Result

  Protected Overridable Function GetConnection(ByVal lpScope As OperationScope) As IRepositoryConnection
    Try
      If Me.Parent Is Nothing Then
        Return Nothing
      Else
        Return GetConnection(Me.Parent, lpScope)
      End If


    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Overridable Function GetPrimaryConnection() As IRepositoryConnection
    Try

      Return GetConnection(Me.Scope)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Function ConvertResult(ByVal lpResult As Boolean) As Result
    Try
      If lpResult = True Then
        Return OperationEnumerations.Result.Success
      Else
        Return OperationEnumerations.Result.Failed
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  'Protected Function GetPathFactory(ByVal lpConnection As IRepositoryConnection, _
  '                              ByVal lpDocument As Document) As PathFactory

  '  Try

  '    Dim lobjBaseFolderProperty As ECMProperty = lpDocument.GetFolderPathProperty()
  '    Dim lstrBaseFolderPath As String = String.Empty
  '    Dim lenuFilingMode As Core.FilingMode = Me.WorkItem.DocumentFilingMode
  '    Dim lobjPathFactory As PathFactory = Nothing

  '    If (lobjBaseFolderProperty Is Nothing) Then

  '      If (Me.Batch.LeadingDelimiter = True) Then
  '        lstrBaseFolderPath = Me.Batch.FolderDelimiter

  '      Else
  '        lstrBaseFolderPath = String.Empty
  '      End If

  '      lenuFilingMode = FilingMode.UnFiled

  '    Else

  '      If (lobjBaseFolderProperty.Values.Count > 0) Then
  '        lstrBaseFolderPath = lobjBaseFolderProperty.Values(0)

  '      Else

  '        If (Me.Batch.LeadingDelimiter = True) Then
  '          lstrBaseFolderPath = Me.Batch.FolderDelimiter

  '        Else
  '          lstrBaseFolderPath = String.Empty
  '        End If

  '        lenuFilingMode = FilingMode.UnFiled
  '      End If

  '    End If

  '    Dim lstrOriginalFolderPath As String = lstrBaseFolderPath

  '    If lpContentSource.ProviderName = "File System Provider" Then
  '      ' This is a file system provider, we need to keep the drive information
  '      lobjPathFactory = New PathFactory(lstrOriginalFolderPath, lstrBaseFolderPath, Me.Batch.BasePathLocation, Me.Batch.FolderDelimiter, False, lenuFilingMode, True)

  '    Else
  '      ' This is not a file system provider, we need to discard the drive information
  '      lobjPathFactory = New PathFactory(lstrOriginalFolderPath, lstrBaseFolderPath, Me.Batch.BasePathLocation, Me.Batch.FolderDelimiter, Me.Batch.LeadingDelimiter, lenuFilingMode, False)
  '    End If

  '    'lobjPathFactory.CreateFolderPath()

  '    Return lobjPathFactory

  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    '  Re-throw the exception to the caller
  '    Throw
  '  End Try

  'End Function

  Protected Overridable Function GetDefaultParameters() As IParameters
    Try
      Return New Parameters
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function GetBooleanParameterValue(ByVal lpOperable As IOperable, ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Boolean
    Try
      Return ActionItem.GetBooleanParameterValue(lpOperable, lpParameterName, lpDefaultValue)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function GetStringParameterValue(ByVal lpOperable As IOperable, ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Boolean
    Try
      Return ActionItem.GetParameterValue(lpOperable, lpParameterName, lpDefaultValue).ToString
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  'Public Shared Function GetParameterValue(ByVal lpOperable As IOperable, ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Object
  '  Try
  '    If lpOperable.Parameters.Contains(lpParameterName) Then
  '      ' Return lpOperable.Parameters.Item(lpParameterName).Value
  '      Dim lobjValue As Object = lpOperable.Parameters.Item(lpParameterName).Value
  '      If lobjValue IsNot Nothing AndAlso TypeOf lobjValue Is String Then
  '        Dim lstrValue As String = lobjValue.ToString
  '        Dim lobjParamNameValuePair As INameValuePair = Process.GetInlineParameter(lstrValue)

  '        If lobjParamNameValuePair Is Nothing Then
  '          Return lstrValue
  '        End If

  '        Dim lstrRequestedParameter As String = String.Empty

  '        Select Case lobjParamNameValuePair.Name.ToLower
  '          Case "parameter"
  '            lstrRequestedParameter = lpOperable.GetStringParameterValue(lobjParamNameValuePair.Value, String.Empty)
  '          Case "processparameter"
  '            If lpOperable.WorkItem IsNot Nothing AndAlso lpOperable.WorkItem.Process IsNot Nothing Then
  '              lstrRequestedParameter = lpOperable.WorkItem.Process.GetStringParameterValue(lobjParamNameValuePair.Value, String.Empty)
  '            Else
  '              Return lstrValue
  '            End If
  '          Case "workitem"
  '            ' Work on this case
  '            Select Case lobjParamNameValuePair.Value.ToLower
  '              Case "sourcedocid"
  '                lstrRequestedParameter = lpOperable.WorkItem.SourceDocId
  '              Case "destdocid", "destinationdocid"
  '                lstrRequestedParameter = lpOperable.WorkItem.DestinationDocId
  '              Case Else
  '                Return lstrValue
  '            End Select
  '        End Select

  '        If String.IsNullOrEmpty(lstrRequestedParameter) Then
  '          Return lstrValue
  '        End If

  '        Dim lobjRegex As Regex = New Regex(mstrParameterExpression, _
  '            RegexOptions.CultureInvariant Or RegexOptions.Compiled)

  '        ' Split the InputText wherever the regex matches
  '        Dim lstrResults As String() = lobjRegex.Split(lstrValue)

  '        ' Test to see if there is a match in the InputText
  '        Dim lblnIsMatch As Boolean = lobjRegex.IsMatch(lstrValue)

  '        If lblnIsMatch Then
  '          Dim lintPrefixGroupNumber As Integer = lobjRegex.GroupNumberFromName("Prefix")
  '          Dim lintSuffixGroupNumber As Integer = lobjRegex.GroupNumberFromName("Suffix")
  '          Dim lobjStringBuilder As New StringBuilder

  '          If lintPrefixGroupNumber > 0 Then
  '            lobjStringBuilder.Append(lstrResults(lintPrefixGroupNumber))
  '          End If
  '          lobjStringBuilder.Append(lstrRequestedParameter)
  '          If lintSuffixGroupNumber > 0 Then
  '            lobjStringBuilder.Append(lstrResults(lintSuffixGroupNumber))
  '          End If

  '          Return lobjStringBuilder.ToString

  '        Else
  '          Return lstrValue
  '        End If
  '      Else
  '        Return lobjValue
  '      End If
  '    Else
  '      Return lpDefaultValue
  '    End If
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  Protected Overridable Function GetBooleanParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Boolean Implements IOperable.GetBooleanParameterValue
    Try
      Return ActionItem.GetBooleanParameterValue(Me, lpParameterName, lpDefaultValue)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Overridable Function GetEnumParameterValue(ByVal lpParameterName As String, ByVal lpEnumType As Type, ByVal lpDefaultValue As Object) As [Enum] Implements IOperable.GetEnumParameterValue
    Try
      Return ActionItem.GetEnumParameterValue(Me, lpParameterName, lpEnumType, lpDefaultValue)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Overridable Function GetStringParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As String Implements IOperable.GetStringParameterValue
    Try
      Return ActionItem.GetStringParameterValue(Me, lpParameterName, lpDefaultValue)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Overridable Function GetParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Object Implements IOperable.GetParameterValue
    Try
      Return ActionItem.GetParameterValue(Me, lpParameterName, lpDefaultValue)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Overridable Sub SetParameterValue(ByVal lpParameterName As String, ByVal lpValue As Object)
    Try
      If Parameters.Contains(lpParameterName) Then
        Parameters.Item(lpParameterName).Value = lpValue
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  'Public Function ResolveInlineParameter(ByVal lpValue As String) As String Implements IOperable.ResolveInlineParameter
  '  Try
  '    Return Operation.ResolveInlineParameter(Me,lpValue)
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  ' ''' <summary>
  ' '''  <para>Resolves the specified process parameter and returns a string with the requested value.</para>
  ' ''' </summary>
  ' ''' <param name="lpValue">
  ' '''  <para>The value to resolve.</para>
  ' ''' </param>
  ' ''' <returns>
  ' '''  <para>Takes an incoming value with an inline parameter reference and resolves the value. </para>
  ' '''  <para>For example, if the incoming value of is 'Export - {ProcessParameter:IdFileBase}' the method will look for a process parameter named
  ' ''' 'IdFileBase'.  If it finds it the method will strip off the section with the curly braces and replace it with the parameter value.  So if the value
  ' ''' of the process parameter <em>IdFileBase</em> is 'CSVExport1' then the return value would be 'Export - CSVExport1'.</para>
  ' ''' </returns>
  'Public Shared Function ResolveInlineParameter(ByVal lpOperable As IOperable, ByVal lpValue As String) As String
  '  Try

  '    Dim lobjParamNameValuePair As INameValuePair = Process.GetInlineParameter(lpValue)

  '    If lobjParamNameValuePair Is Nothing Then
  '      Return lpValue
  '    End If

  '    Dim lstrRequestedParameter As String = String.Empty

  '    Select Case lobjParamNameValuePair.Name.ToLower
  '      Case "parameter"
  '        lstrRequestedParameter = lpOperable.GetParameterValue(lobjParamNameValuePair.Value, String.Empty)
  '      Case "processparameter"
  '        If lpOperable.WorkItem IsNot Nothing AndAlso lpOperable.WorkItem.Process IsNot Nothing Then
  '          lstrRequestedParameter = lpOperable.WorkItem.Process.GetParameterValue(lobjParamNameValuePair.Value, String.Empty)
  '        Else
  '          Return lpValue
  '        End If
  '      Case "workitem"
  '        ' Work on this case
  '        Select Case lobjParamNameValuePair.Value.ToLower
  '          Case "sourcedocid"
  '            lstrRequestedParameter = lpOperable.WorkItem.SourceDocId
  '          Case "destdocid", "destinationdocid"
  '            lstrRequestedParameter = lpOperable.WorkItem.DestinationDocId
  '          Case Else
  '            Return lpValue
  '        End Select
  '    End Select

  '    If String.IsNullOrEmpty(lstrRequestedParameter) Then
  '      Return lpValue
  '    End If

  '    Dim lobjRegex As Regex = New Regex(mstrParameterExpression, _
  '        RegexOptions.CultureInvariant Or RegexOptions.Compiled)

  '    ' Split the InputText wherever the regex matches
  '    Dim lstrResults As String() = lobjRegex.Split(lpValue)

  '    ' Test to see if there is a match in the InputText
  '    Dim lblnIsMatch As Boolean = lobjRegex.IsMatch(lpValue)

  '    If lblnIsMatch Then
  '      Dim lintPrefixGroupNumber As Integer = lobjRegex.GroupNumberFromName("Prefix")
  '      Dim lintSuffixGroupNumber As Integer = lobjRegex.GroupNumberFromName("Suffix")
  '      Dim lobjStringBuilder As New StringBuilder

  '      If lintPrefixGroupNumber > 0 Then
  '        lobjStringBuilder.Append(lstrResults(lintPrefixGroupNumber))
  '      End If
  '      lobjStringBuilder.Append(lstrRequestedParameter)
  '      If lintSuffixGroupNumber > 0 Then
  '        lobjStringBuilder.Append(lstrResults(lintSuffixGroupNumber))
  '      End If

  '      Return lobjStringBuilder.ToString

  '    Else
  '      Return lpValue
  '    End If

  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  Protected Overridable Sub RunPreOperationChecks(ByVal lpIncludeAttachmentCheck As Boolean)
    Try

      If Me.WorkItem Is Nothing Then
        Throw New OperationNotInitializedException(Me, "WorkItem not initialized")
      End If

      If Me.Parent Is Nothing Then
        Throw New OperationNotInitializedException(Me, "Parent not initialized")
      End If

      If PrimaryConnection Is Nothing Then
        Throw New OperationNotInitializedException(Me, String.Format("Parent {0} connection not initialized", Me.Scope.ToString))
      End If

      If lpIncludeAttachmentCheck = True Then
        ' Make sure we have a document reference
        If Me.WorkItem.Document Is Nothing Then
          If ((Not String.IsNullOrEmpty(Me.WorkItem.SourceDocId) AndAlso
               (Me.WorkItem.SourceDocId.ToLower.EndsWith(".cpf")) AndAlso
               (IO.File.Exists(Me.WorkItem.SourceDocId)))) Then
            Me.WorkItem.Document = New Document(Me.WorkItem.SourceDocId)
          Else
            Throw New DocumentReferenceNotSetException("Cannot initialize source document, no Document reference set in the work item.")
          End If
        End If
      End If

      If Me.Scope = OperationScope.Source Then
        Me.DocumentId = Me.WorkItem.SourceDocId
      Else
        Me.DocumentId = Me.WorkItem.DestinationDocId
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overridable Sub RunPreOperationChecksForFolder(ByVal lpIncludeAttachmentCheck As Boolean)
    Try

      If Me.WorkItem Is Nothing Then
        Throw New OperationNotInitializedException(Me, "WorkItem not initialized")
      End If

      If Me.Parent Is Nothing Then
        Throw New OperationNotInitializedException(Me, "Parent not initialized")
      End If

      If PrimaryConnection Is Nothing Then
        Throw New OperationNotInitializedException(Me, String.Format("Parent {0} connection not initialized", Me.Scope.ToString))
      End If

      If lpIncludeAttachmentCheck = True Then
        ' Make sure we have a folder reference
        If Me.WorkItem.Folder Is Nothing Then
          If ((Not String.IsNullOrEmpty(Me.WorkItem.SourceDocId) AndAlso
               (Me.WorkItem.SourceDocId.ToLower.EndsWith(".cff")) AndAlso
               (IO.File.Exists(Me.WorkItem.SourceDocId)))) Then
            Me.WorkItem.Folder = New Folder(Me.WorkItem.SourceDocId)
          Else
            Throw New DocumentReferenceNotSetException("Cannot initialize source folder, no Folder reference set in the work item.")
          End If
        End If
      End If

      If Me.Scope = OperationScope.Source Then
        Me.DocumentId = Me.WorkItem.SourceDocId
      Else
        Me.DocumentId = Me.WorkItem.DestinationDocId
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overridable Sub RunPreOperationChecksForObject(ByVal lpIncludeAttachmentCheck As Boolean)
    Try

      If Me.WorkItem Is Nothing Then
        Throw New OperationNotInitializedException(Me, "WorkItem not initialized")
      End If

      If Me.Parent Is Nothing Then
        Throw New OperationNotInitializedException(Me, "Parent not initialized")
      End If

      If PrimaryConnection Is Nothing Then
        Throw New OperationNotInitializedException(Me, String.Format("Parent {0} connection not initialized", Me.Scope.ToString))
      End If

      If lpIncludeAttachmentCheck = True Then
        ' Make sure we have an object reference
        If Me.WorkItem.Object Is Nothing Then
          If ((Not String.IsNullOrEmpty(Me.WorkItem.SourceDocId) AndAlso
               (Me.WorkItem.SourceDocId.ToLower.EndsWith(".cof")) AndAlso
               (IO.File.Exists(Me.WorkItem.SourceDocId)))) Then
            Me.WorkItem.Object = New CustomObject(Me.WorkItem.SourceDocId)
          Else
            Throw New DocumentReferenceNotSetException("Cannot initialize source object, no CustomObject reference set in the work item.")
          End If
        End If
      End If

      If Me.Scope = OperationScope.Source Then
        Me.DocumentId = Me.WorkItem.SourceDocId
      Else
        Me.DocumentId = Me.WorkItem.DestinationDocId
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Sub UpdateParameterToEnum(lpName As String, lpEnumType As Type) Implements IOperable.UpdateParameterToEnum
    Try
      ParameterFactory.UpdateParameterToEnum(Me.Parameters, lpName, lpEnumType)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Methods"

  Public Sub SetInstanceId(lpInstanceId As String) Implements IOperable.SetInstanceId
    Try
      mstrInstanceId = lpInstanceId
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  ''' <summary>
  ''' Called to execute the operation.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Execute(ByVal lpWorkItem As IWorkItem) As OperationEnumerations.Result Implements IOperation.Execute
    Try

      'LogSession.EnterMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))

      ' If applicable, run any pre-operation(s)
      If RunBeforeBegin IsNot Nothing Then
        Dim lenuRunBeforeBeginResult As Result = RunBeforeBegin.Execute(lpWorkItem)
        If lenuRunBeforeBeginResult = OperationEnumerations.Result.Failed Then
          OnError(New OperableErrorEventArgs(Me, lpWorkItem, lpWorkItem.ProcessedMessage))
        End If
      End If

      OnBegin(New OperableEventArgs(Me, lpWorkItem))

      menuResult = OnExecute()

      If (menuResult = OperationEnumerations.Result.Failed) Then
        OnError(New OperableErrorEventArgs(Me, lpWorkItem, lpWorkItem.ProcessedMessage))
        'LogSession.LogDebug("Execution of {0} failed.", Me.Name)
        'LogSession.LogObject(Level.Debug, lpWorkItem)
      Else
        'LogSession.LogDebug("Execution of {0} succeeded.", Me.Name)
      End If

      'If (mblnLogResult) Then
      '  OnComplete(New OperableEventArgs(Me, lpBatchItem))
      'End If

      OnComplete(New OperableEventArgs(Me, lpWorkItem))

      ' If applicable, run any post-operation(s)
      If RunAfterComplete IsNot Nothing Then
        Dim lenuRunAfterCompleteResult As Result = RunAfterComplete.Execute(lpWorkItem)
        If lenuRunAfterCompleteResult = OperationEnumerations.Result.Failed Then
          OnError(New OperableErrorEventArgs(Me, lpWorkItem, lpWorkItem.ProcessedMessage))
        End If
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      OnError(New OperableErrorEventArgs(Me, lpWorkItem, ex))

      ' If applicable, run any post-error operation(s)
      If RunOnFailure IsNot Nothing Then
        Dim lenuRunOnFailureResult As Result = RunOnFailure.Execute(lpWorkItem)
        If lenuRunOnFailureResult = OperationEnumerations.Result.Failed Then
          OnError(New OperableErrorEventArgs(Me, lpWorkItem, lpWorkItem.ProcessedMessage))
        End If
      End If
    Finally
      'LogSession.LeaveMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
    End Try

    Return menuResult

  End Function

  Public Function Rollback(ByVal lpWorkItem As IWorkItem) As Result Implements IOperation.Rollback
    Try
      Return OnRollback()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  'Public Function Rollback() As Result Implements IOperable.Rollback
  '	Try
  '		Return OnRollback()
  '	Catch ex As Exception
  '		ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '		'   Re-throw the exception to the caller
  '		Throw
  '	End Try
  'End Function

  ''' <summary>Resets the operation.</summary>
  Public Overridable Sub Reset() Implements IOperable.Reset
    Try
      menuResult = Result.NotProcessed
      mstrProcessedMessage = String.Empty
      mobjResultDetail = Nothing
      mdatStartTime = DateTime.MinValue
      mdatFinishTime = DateTime.MinValue
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  ''' <summary>
  ''' To be called as the first part of the Execute method for any operation.
  ''' </summary>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Public Overridable Sub OnBegin(ByVal e As OperableEventArgs) Implements IOperable.OnBegin
    Try

      'ApplicationLogging.WriteLogEntry(String.Format("================Operation Begin: {0}, {1}, {2}", e.DocumentId, e.Operation.Name, e.WorkItem.SourceDocId), TraceEventType.Information, 43440)


      ' Since all the Operation subclasses should call this method before executing
      ' we can ensure the initialization of the WorkItem and Parent properties here.
      WorkItem = e.WorkItem

      ' Set the start time
      StartTime = Now

      If Parent Is Nothing AndAlso WorkItem.Parent IsNot Nothing Then
        Parent = WorkItem.Parent
      End If

      ' Raise the Begin event
      RaiseEvent Begin(Me, e)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overridable Sub OnComplete(ByVal e As OperableEventArgs) Implements IOperable.OnComplete
    Try

      'ApplicationLogging.WriteLogEntry(String.Format("================Operation End: {0}, {1}, {2}", e.DocumentId, e.Operation.Name, e.WorkItem.SourceDocId), TraceEventType.Information, 43441)
      ' 'LogSession.EnterMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
      ' Stop the clock
      FinishTime = Now

      If Not String.IsNullOrEmpty(Me.ProcessedMessage) AndAlso String.IsNullOrEmpty(e.WorkItem.ProcessedMessage) Then
        e.WorkItem.ProcessedMessage = Me.ProcessedMessage
      End If

      ' Initialize the detailed results
      mobjResultDetail = New OperationResult(Me)

      If Not String.IsNullOrEmpty(Me.ProcessedMessage) Then
        'LogSession.LogMessage(Me.ProcessedMessage)
      End If

      RaiseEvent Complete(Me, e)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    Finally
      ' 'LogSession.LeaveMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
    End Try
  End Sub

  Public Overridable Sub OnError(ByVal e As OperableErrorEventArgs) Implements IOperable.OnError
    Try

      'ApplicationLogging.WriteLogEntry(String.Format("================Operation Error: {0}, {1}, {2}", e.DocumentId, e.Operation.Name, e.WorkItem.SourceDocId), TraceEventType.Information, 43442)

      ' Stop the clock
      FinishTime = Now

      menuResult = OperationEnumerations.Result.Failed

      If String.IsNullOrEmpty(Me.ProcessedMessage) Then
        If e.Exception IsNot Nothing Then
          Me.ProcessedMessage = $"{Me.DisplayName}: {e.Exception.Message}"
        Else
          Me.ProcessedMessage = e.Message
          If String.IsNullOrEmpty(Me.WorkItem.ProcessedMessage) Then
            Me.WorkItem.ProcessedMessage = e.Message
          End If
        End If
      End If

      'LogSession.LogError(Me.ProcessedMessage)

      RaiseEvent OperatingError(Me, e)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub SetDescription(ByVal lpDescription As String) Implements IOperable.SetDescription ', IActionItem.SetDescription
    Try
      mstrDescription = lpDescription
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub SetResult(ByVal lpResult As Result) Implements IOperation.SetResult
    Try
      menuResult = lpResult
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Function ToActionItem() As IActionItem
    Try
      Return New ActionItem(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Function ToString() As String Implements IOperable.ToString
    Try
      Return DebuggerIdentifier()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function ToXmlItemString() As String Implements IActionItem.ToXmlElementString
    Try
      Return Serializer.Serialize.XmlElementString(Me.ToActionItem())
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function ToXmlString() As String Implements IOperable.ToXmlElementString
    Try
      Return Serializer.Serialize.XmlElementString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Friend Methods"

  ''' <summary>
  ''' Implements the operation execution
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend MustOverride Function OnExecute() As OperationEnumerations.Result Implements IOperation.OnExecute

  Friend Overridable Function OnRollback() As Result Implements IOperation.OnRollBack
    Try
      If CanRollback = False Then
        Return OperationEnumerations.Result.RollbackNotSupported
      Else
        Return OperationEnumerations.Result.RollbackNotImplemented
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Shared Function GetConnection(ByVal lpParent As Object, ByVal lpScope As OperationScope) As IRepositoryConnection
    Try
      Dim lobjConnection As IRepositoryConnection

      If lpScope = OperationScope.Source Then
        lobjConnection = lpParent.SourceConnection
      Else
        lobjConnection = lpParent.DestinationConnection
      End If

      If lobjConnection IsNot Nothing Then
        If lobjConnection.IsDisposed Then
          Select Case lpScope
            Case OperationScope.Source
              lpParent.RefreshSourceConnection()
              lobjConnection = lpParent.SourceConnection
            Case OperationScope.Destination
              lpParent.RefreshDestinationConnection()
              lobjConnection = lpParent.DestinationConnection
          End Select
        End If
        lobjConnection.Provider.Connect(lobjConnection)
        lpParent.ExportPath = lobjConnection.ExportPath
        lobjConnection.Provider.Tag = lpParent.Id
      End If

      Return lobjConnection

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "IJsonSerializable(Of IOperation)"

  Public Overloads Function ToJson() As String Implements IJsonSerializable(Of IOperation).ToJson, IOperable.ToJson
    Try
      Return WriteJsonString()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function FromJson(lpJson As String) As IOperation Implements IJsonSerializable(Of IOperation).FromJson
    Try
      Return JsonConvert.DeserializeObject(lpJson, GetType(IOperation), DefaultJsonSerializerSettings.Settings)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Function WriteJsonString() As String
    Try
      Return JsonConvert.SerializeObject(Me, Newtonsoft.Json.Formatting.None, New OperationJsonConverter())
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function CreateFromJsonReader(reader As JsonReader) As IOperation
    Try
      'Return JsonConvert.DeserializeObject(reader, GetType(IOperation), New OperationJsonConverter())
      Dim lobjConverter As New OperationJsonConverter()
      Return lobjConverter.ReadJson(reader, Nothing, Nothing, Nothing)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function CreateFromJson(lpJson As String) As IOperation
    Try
      Return JsonConvert.DeserializeObject(lpJson, GetType(IOperation), New OperationJsonConverter())
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "IXmlSerializable Implementation"

  Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
    Try
      Return Nothing
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Sub ReadOperableXml(ByRef sender As IOperable, ByVal reader As System.Xml.XmlReader)

    Try

      If reader.IsEmptyElement Then
        reader.Read()
        Exit Sub
      End If

      Dim lstrCurrentElementName As String = String.Empty

      sender.SetDescription(reader.GetAttribute("Description"))

      Boolean.TryParse(reader.GetAttribute("LogResult"), sender.LogResult)

      If TypeOf sender Is IOperation Then
        Dim lstrScope As String = reader.GetAttribute("Scope")
        If Not String.IsNullOrEmpty(lstrScope) Then
          CType(sender, IOperation).Scope = CType([Enum].Parse(GetType(OperationScope), reader.GetAttribute("Scope")), OperationScope)
        End If
      End If

      Do Until reader.NodeType = XmlNodeType.EndElement AndAlso (reader.Name.EndsWith("Operation") OrElse reader.Name = "Process")
        If reader.NodeType = XmlNodeType.Element Then
          lstrCurrentElementName = reader.Name
        Else
          Select Case lstrCurrentElementName
            Case "Parameters"
              ' Skip to the next line
            Case "Parameter"
              ' TODO: Add the parameter
            Case "Values"
              ' Skip to the next line
            Case "Value"

            Case "TrueOperations"
              ' TODO: Add the TrueOperations

            Case "FalseOperations"
              ' TODO: Add the FalseOperations

            Case "RunBeforeBegin"
              ' TODO: Implement this...

            Case "RunAfterComplete"
              ' TODO: Implement this...

            Case "RunOnFailure"
              ' TODO: Implement this...

          End Select
        End If
      Loop

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
    Try
      ReadOperableXml(Me, reader)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Shared Sub WriteOperableXml(ByVal sender As IOperable, ByVal writer As System.Xml.XmlWriter)
    Try

      With writer

        ' Write the Name attribute
        .WriteAttributeString("Name", sender.Name)

        ' Write the Description attribute
        .WriteAttributeString("Description", sender.Description)

        ' Write the LogResult attribute
        .WriteAttributeString("LogResult", sender.LogResult.ToString)

        If TypeOf sender Is IOperation Then
          ' Write the scope
          .WriteAttributeString("Scope", CType(sender, IOperation).Scope.ToString)
        End If

        ' Write the result, if applicable
        If sender.Result <> OperationEnumerations.Result.NotProcessed AndAlso sender.LogResult = True Then
          .WriteAttributeString("Result", sender.Result.ToString)
        End If

        ' Write the result, if applicable
        If sender.Result <> OperationEnumerations.Result.NotProcessed AndAlso sender.LogResult = True Then
          .WriteAttributeString("ProcessedMessage", sender.ProcessedMessage)
        End If

        ' Write the times, if applicable
        If sender.StartTime <> DateTime.MinValue AndAlso sender.LogResult = True Then
          .WriteAttributeString("StartTime", Helper.ToDetailedDateString(sender.StartTime, sender.Locale.Name))
        End If

        If sender.FinishTime <> DateTime.MinValue AndAlso sender.LogResult = True Then
          .WriteAttributeString("FinishTime", Helper.ToDetailedDateString(sender.FinishTime, sender.Locale.Name))
        End If

        If sender.TotalProcessingTime <> TimeSpan.Zero AndAlso sender.LogResult = True Then
          .WriteAttributeString("TotalProcessingTime", sender.TotalProcessingTime.ToString)
        End If

        ' Write the Parameters
        ' Open the Parameters Element
        .WriteStartElement("Parameters")

        If sender.Parameters IsNot Nothing Then
          For Each lobjParameter As IParameter In sender.Parameters
            ' Write the Parameter element
            .WriteRaw(lobjParameter.ToXmlString)
          Next
        End If

        ' End the Parameters element
        .WriteEndElement()

        ' Write the RunBeforeBegin
        ' Open the RunBeforeBegin Element
        .WriteStartElement("RunBeforeBegin")

        If sender.RunBeforeBegin IsNot Nothing Then
          ' Write the operable element
          .WriteRaw(sender.RunBeforeBegin.ToXmlElementString)
        End If

        ' End the RunBeforeBegin element
        .WriteEndElement()

        ' Write the RunAfterComplete
        ' Open the RunAfterComplete Element
        .WriteStartElement("RunAfterComplete")

        If sender.RunAfterComplete IsNot Nothing Then
          ' Write the operable element
          .WriteRaw(sender.RunAfterComplete.ToXmlElementString)
        End If

        ' End the RunAfterComplete element
        .WriteEndElement()

        ' Write the RunOnFailure
        ' Open the RunOnFailure Element
        .WriteStartElement("RunOnFailure")

        If sender.RunOnFailure IsNot Nothing Then
          ' Write the operable element
          .WriteRaw(sender.RunOnFailure.ToXmlElementString)
        End If

        ' End the RunOnFailure element
        .WriteEndElement()

      End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overridable Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
    Try
      WriteOperableXml(Me, writer)
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
      Return OperationFactory.Create(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "IDisposable Implementation"

  Private disposedValue As Boolean = False    ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(ByVal disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        'FinalizeLogSession()
        ' DISPOSETODO: free other state (managed objects).
        mstrDocumentId = Nothing
        mstrProcessedMessage = Nothing
        menuResult = Nothing
        mblnLogResult = Nothing
        If mobjResultDetail IsNot Nothing Then
          mobjResultDetail.Dispose()
          mobjResultDetail = Nothing
        End If
        menuScope = Nothing
        mobjParameters = Nothing
        mstrDescription = Nothing
        mobjParent = Nothing
        mobjWorkItem = Nothing
        mdatStartTime = Nothing
        mdatFinishTime = Nothing
        mobjSourceConnection = Nothing
        mobjDestinationConnection = Nothing
        mobjPrimaryConnection = Nothing
      End If

      ' DISPOSETODO: free your own state (unmanaged objects).
      ' DISPOSETODO: set large fields to null.
    End If
    Me.disposedValue = True
  End Sub

#Region " IDisposable Support "
  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    Dispose(True)
    GC.SuppressFinalize(Me)
  End Sub

#End Region

#End Region

  Public Overridable Sub CheckParameters() Implements IOperable.CheckParameters
    Try
      ' Do Nothing
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub mobjParameters_ItemPropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles mobjParameters.ItemPropertyChanged
    RaiseEvent ParameterPropertyChanged(sender, e)
  End Sub

  'Private Sub mobjParameters_ItemPropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles mobjParameters.ItemPropertyChanged
  '  Beep()
  'End Sub
End Class