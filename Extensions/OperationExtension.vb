' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperationExtension.vb
'  Description :  Base class from which all operation extensions should inherit.
'  Created     :  11/16/2011 10:14:30 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Globalization
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Extensions
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json


#End Region

Namespace Extensions

  Public MustInherit Class OperationExtension
    Inherits Extension
    Implements IOperationExtension
    Implements IOperationInformation
    Implements IXmlSerializable
    Implements ICloneable
    Implements IExtensionInformation
    'Implements ILoggable
    ' Implements IWorkItem

#Region "Class Constants"

    Private Const IOPERATIONEXTENSION_NAME As String = "IOperationExtension"

    Protected Shadows NAME_CONST As String = "Name"
    Protected Shadows DESC_CONST As String = "Description"
    Protected OPERATION_CONST As String = "Operation"
    Friend Const LOG_RESULT As String = "LogResult"

#End Region

#Region "Class Variables"

    Private mstrProcessedMessage As String = String.Empty
    'Private mobjFailureOperations As New Operations
    Private WithEvents MobjParent As IItemParent
    Private mobjWorkItem As IWorkItem = Nothing

    Private mstrDocumentId As String = String.Empty
    Private WithEvents MobjParameters As New Parameters
    Protected menuResult As OperationEnumerations.Result = OperationEnumerations.Result.NotProcessed
    Private mobjResultDetail As IOperableResult = Nothing
    Private mblnLogResult As Boolean = True
    Private menuScope As OperationScope = OperationScope.Source
    Private mstrExtensionPath As String = String.Empty
    Private mdatStartTime As DateTime = DateTime.MinValue
    Private mdatFinishTime As DateTime = DateTime.MinValue

    Private mobjSourceConnection As IRepositoryConnection = Nothing
    Private mobjDestinationConnection As IRepositoryConnection = Nothing
    Private mobjPrimaryConnection As IRepositoryConnection = Nothing
    Private ReadOnly mobjLocale As CultureInfo = CultureInfo.CurrentCulture
    'Private mstrLocale As String = mobjLocale.Name
    Private mobjHost As Object = Nothing
    Private mobjOperableParent As IOperable = Nothing
    Private mobjTag As Object = Nothing
    Private mobjRunBeforeBegin As IOperable = Nothing
    Private mobjRunAfterComplete As IOperable = Nothing
    Private mobjRunOnFailure As IOperable = Nothing
    Private mstrInstanceId As String = String.Empty

    'Private mobjLogSession As Gurock.SmartInspect.Session

#End Region

#Region "Public Properties"

    Public Property OperableParent As IOperable Implements IOperationExtension.OperableParent
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

    Public Property Parent As IItemParent Implements IOperationExtension.Parent
      Get
        Return MobjParent
      End Get
      Set(ByVal value As IItemParent)
        MobjParent = value
      End Set
    End Property

    Public Property WorkItem As IWorkItem Implements IOperationExtension.WorkItem
      Get
        Return mobjWorkItem
      End Get
      Set(ByVal value As IWorkItem)
        mobjWorkItem = value
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
            WorkItem.ProcessedMessage = value
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overridable Shadows ReadOnly Property Description As String Implements IOperable.Description, IOperationInformation.Description
      Get
        Return MyBase.Description
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

    'Public Overrides ReadOnly Property Type As ProjectExtensionType
    '  Get
    '    Return ProjectExtensionType.Operation
    '  End Get
    'End Property

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

    <Category(Operation.CATEGORY_BEHAVIOR)>
    Public Property RunBeforeBegin As IOperable Implements IOperable.RunBeforeBegin
      Get
        Return mobjRunBeforeBegin
      End Get
      Set(value As IOperable)
        mobjRunBeforeBegin = value
      End Set
    End Property

    <Category(Operation.CATEGORY_BEHAVIOR)>
    Public Property RunAfterComplete As IOperable Implements IOperable.RunAfterComplete
      Get
        Return mobjRunAfterComplete
      End Get
      Set(value As IOperable)
        mobjRunAfterComplete = value
      End Set
    End Property

    <Category(Operation.CATEGORY_BEHAVIOR)>
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

    Public MustOverride ReadOnly Property CanRollback As Boolean Implements IOperable.CanRollback

#End Region

#Region "Constructors"

    ''' <summary>
    ''' The default constructor is not public.  All construction of 
    ''' extensions should go throw the CreateExtension method.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub New()
      MyBase.New()
      Try
        If Parameters.Count = 0 Then
          Parameters = GetDefaultParameters()
        End If
        'SetDisplayName(DisplayName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Shared Methods"

    Public Sub SetInstanceId(lpInstanceId As String) Implements IOperable.SetInstanceId
      Try
        mstrInstanceId = lpInstanceId
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Function LoadExtensions(ByVal lpExtensionPath As String) As IOperationExtensions
      Try

        Dim lobjAssembly As System.Reflection.Assembly
        Dim lobjExtensionCandidate As Type
        Dim lobjOperationExtension As OperationExtension
        Dim lobjOperationExtensions As New OperationExtensions

        If String.IsNullOrEmpty(lpExtensionPath) Then
          Throw New ArgumentNullException(NameOf(lpExtensionPath))
        End If

        If IO.File.Exists(lpExtensionPath) = False Then
          Throw New InvalidPathException(
            String.Format("There is no extension file in the specified path '{0}'.",
                          lpExtensionPath), lpExtensionPath)
        End If

        If lpExtensionPath.ToLower.EndsWith(".dll") = False Then
          Throw New InvalidExtensionException(String.Format("The file '{0}' is not a dll file.",
                                                            IO.Path.GetFileName(lpExtensionPath)))
        End If

        Try
          lobjAssembly = System.Reflection.Assembly.LoadFrom(lpExtensionPath)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 64339, lpExtensionPath)
          ' If we are unable to load the assembly then it is not valid.
          Return Nothing
        End Try

        Dim lobjTypes As Type() = lobjAssembly.GetTypes

        For Each lobjType As Type In lobjTypes
          lobjExtensionCandidate = lobjType.GetInterface(IOPERATIONEXTENSION_NAME)
          If lobjExtensionCandidate IsNot Nothing AndAlso lobjType.IsAbstract = False Then
            lobjOperationExtension = CType(lobjAssembly.CreateInstance(lobjType.FullName), OperationExtension)
            lobjOperationExtensions.Add(lobjOperationExtension)
          End If
        Next

        Return lobjOperationExtensions

        Throw New InvalidExtensionException(String.Format("The assembly '{0}' does not implement IOperationExtension.",
                                                          IO.Path.GetFileName(lpExtensionPath)))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 64253, lpExtensionPath)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Overrides Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New Text.StringBuilder
      Try

        If disposedValue = True Then
          If String.IsNullOrEmpty(Me.Name) Then
            Return "Operation Extension Disposed"
          Else
            Return String.Format("{0} Operation Extension Disposed", Me.Name)
          End If
        End If

        lobjIdentifierBuilder.AppendFormat("{0} Operation Extension", Me.GetType.Name)

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Function GetOperationName() As String
      Try
        Return Me.GetType.Name.Replace("Operation", String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Function GetDefaultParameters() As IParameters
      Try
        Return New Parameters
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

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
    '    Return Operation.ResolveInlineParameter(Me, lpValue)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Protected Overridable Function GetConnection(ByVal lpScope As OperationScope) As IRepositoryConnection
      Try
        Return Operation.GetConnection(Me.Parent, lpScope)

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

    Protected Shared Function ConvertResult(ByVal lpResult As Boolean) As Result
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

    Protected Friend ReadOnly Property SourceConnection As IRepositoryConnection
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

    Protected Friend ReadOnly Property DestinationConnection As IRepositoryConnection
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

    Protected Friend ReadOnly Property PrimaryConnection As IRepositoryConnection
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
            Throw New DocumentReferenceNotSetException("Cannot initialize source document, no Document reference set in the work item.")
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

#Region "Private Methods"

    'Protected Overrides Sub InitializeIdentity()
    '  Try
    '    MyBase.InitializeIdentity()
    '    mstrOperationType = OPERATION_CONST
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

#End Region

#Region "IOperationExtension Implementation"

#Region "Public Events"

    Public Event Begin(ByVal sender As Object, ByVal e As OperableEventArgs) Implements IOperation.Begin

    Public Event Complete(ByVal sender As Object, ByVal e As OperableEventArgs) Implements IOperation.Complete

    Public Event OperatingError(ByVal sender As Object, ByVal e As OperableErrorEventArgs) Implements IOperation.OperatingError
    Public Event ParameterPropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements IOperable.ParameterPropertyChanged

#End Region

    Public Property DocumentId As String Implements IOperable.DocumentId
      Get
        Return mstrDocumentId
      End Get
      Set(ByVal value As String)
        mstrDocumentId = value
      End Set
    End Property

    Public Property Parameters As IParameters Implements IOperable.Parameters
      Get
        Return MobjParameters
      End Get
      Set(ByVal value As IParameters)
        MobjParameters = value
      End Set
    End Property

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

    'Public MustOverride ReadOnly Property CanRollback As Boolean

    Public ReadOnly Property Result As OperationEnumerations.Result Implements IOperable.Result
      Get
        Return menuResult
      End Get
    End Property

    Public ReadOnly Property ResultDetail As IOperableResult Implements IOperable.ResultDetail
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
    ''' <remarks>This is primarily for controlling when the operation 
    ''' results should be written to the Job Manager database.</remarks>
    Public Property LogResult As Boolean Implements IOperable.LogResult
      Get
        Return mblnLogResult
      End Get
      Set(ByVal value As Boolean)
        mblnLogResult = value
      End Set
    End Property

    Public MustOverride Overloads ReadOnly Property Name As String Implements IOperation.Name, IOperationInformation.Name

    Public Overrides ReadOnly Property DisplayName As String Implements IOperable.DisplayName, IOperationInformation.DisplayName
      Get
        Return Helper.CreateDisplayName(Me.Name)
      End Get
    End Property

    Public MustOverride Overloads ReadOnly Property CompanyName As String Implements IOperationInformation.CompanyName

    Public MustOverride Overloads ReadOnly Property ProductName As String Implements IOperationInformation.ProductName

    Public Property Scope As OperationEnumerations.OperationScope Implements IOperation.Scope
      Get
        Return menuScope
      End Get
      Set(ByVal value As OperationEnumerations.OperationScope)
        menuScope = value
      End Set
    End Property

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

    Public Property ShouldExecute As Boolean Implements IOperable.ShouldExecute

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
    Public ReadOnly Property TotalOperationProcessingTime As TimeSpan Implements IOperation.TotalProcessingTime
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

    Public Overridable Sub CheckParameters() Implements IOperable.CheckParameters
      Try
        ' Do Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Executes the operation.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Execute(ByRef lpWorkItem As IWorkItem) As OperationEnumerations.Result Implements IOperation.Execute
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
        'RaiseEvent OperatingError(Me, New OperableErrorEventArgs(Me, lpWorkItem, ex))
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

    Public MustOverride Function Rollback() As Result

    Public Overridable Sub Reset() Implements IOperation.Reset
      Try
        menuResult = Result.NotProcessed
        mobjResultDetail = Nothing
        mdatStartTime = DateTime.MinValue
        mdatFinishTime = DateTime.MinValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected MustOverride Function OnExecute() As OperationEnumerations.Result Implements IOperationExtension.OnExecute

    Protected Overridable Function OnRollback() As Result Implements IOperation.OnRollBack
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

    Public Overloads Sub SetDescription(ByVal lpDescription As String) Implements IOperable.SetDescription ', IActionItem.SetDescription
      Try
        MyBase.SetDescription(lpDescription)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub SetResult(ByVal lpResult As Result) Implements IOperationExtension.SetResult
      Try
        menuResult = lpResult
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overridable Sub OnBegin(ByVal e As OperableEventArgs) Implements IOperable.OnBegin
      Try

        ' Set the start time
        StartTime = Now

        ' Since all the Operation subclasses should call this method before executing
        ' we can ensure the initialization of the WorkItem and Parent properties here.
        WorkItem = e.WorkItem
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

        ' Stop the clock
        FinishTime = Now

        If Not String.IsNullOrEmpty(Me.ProcessedMessage) AndAlso String.IsNullOrEmpty(e.WorkItem.ProcessedMessage) Then
          e.WorkItem.ProcessedMessage = Me.ProcessedMessage
        End If

        ' Initialize the detailed results
        mobjResultDetail = New OperationResult(Me)

        RaiseEvent Complete(Me, e)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overridable Sub OnError(ByVal e As OperableErrorEventArgs) Implements IOperable.OnError
      Try

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

        ApplicationLogging.WriteLogEntry(Me.ProcessedMessage, TraceEventType.Error, 65238)
        RaiseEvent OperatingError(Me, e)

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

    Public Function ToJson() As String Implements IOperable.ToJson
      Try
        Return WriteJsonString()
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

    Friend Function WriteJsonString() As String
      Try
        Return JsonConvert.SerializeObject(Me, Newtonsoft.Json.Formatting.None, New OperationJsonConverter())
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IBatchItem Implementation"

    'Public Property BatchId As String Implements IBatchItem.BatchId
    '  Get
    '    Return mstrBatchId
    '  End Get
    '  Set(ByVal value As String)
    '    mstrBatchId = value
    '  End Set
    'End Property

    'Public Sub BeginProcessItem() Implements IBatchItem.BeginProcessItem

    '  Try
    '    mobjBatchItemProcessEventArgs = BatchItemProcessEventArgs.InitializeEvent(Me.Id, Me.BatchId, Me.Title, Me.Operation, Me.Batch.MachineName)
    '    Me.Batch.BatchContainer.BeginProcessItem(mobjBatchItemProcessEventArgs)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try

    'End Sub

    'Public Property DestDocId As String Implements IBatchItem.DestDocId
    '  Get
    '    Return mstrDestDocId
    '  End Get
    '  Set(ByVal value As String)
    '    mstrDestDocId = value
    '  End Set
    'End Property

    'Public Sub EndProcessItem(lpProcessedStatus As ProcessedStatus, _
    '                          lpProcessedMessage As String, _
    '                          lpDestDocId As String) Implements IBatchItem.EndProcessItem

    '  Try
    '    EndProcessItem(lpProcessedStatus, lpProcessedMessage, lpDestDocId, Now)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try

    'End Sub

    'Public Sub EndProcessItem(lpProcessedStatus As ProcessedStatus, _
    '                          lpProcessedMessage As String, _
    '                          lpDestDocId As String, _
    '                          lpEndTime As Date) Implements IBatchItem.EndProcessItem

    '  Try
    '    With mobjBatchItemProcessEventArgs
    '      .ProcessedStatus = lpProcessedStatus
    '      .ProcessedMessage = lpProcessedMessage
    '      .DestDocId = lpDestDocId
    '      .EndTime = lpEndTime
    '    End With

    '    Me.Batch.BatchContainer.EndProcessItem(mobjBatchItemProcessEventArgs)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try

    'End Sub

    'Public Function Execute(lpProcess As IOperable) As Result Implements IOperationExtension.Execute
    '  Try

    '    OnBegin(New OperableEventArgs(Me, lpProcess.BatchItem))

    '    menuResult = OnExecute(lpProcess.BatchItem)

    '    'If (mblnLogResult) Then
    '    '  OnComplete(New OperableEventArgs(Me, lpBatchItem))
    '    'End If

    '    OnComplete(New OperableEventArgs(Me, lpProcess.BatchItem))

    '    Return menuResult

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    OnError(New OperableErrorEventArgs(Me, lpProcess.BatchItem, ex))
    '  End Try
    'End Function

    'Public Overloads Property Id As String Implements IBatchItem.Id
    '  Get
    '    Return MyBase.Id
    '  End Get
    '  Set(value As String)
    '    MyBase.Id = value
    '  End Set
    'End Property

    'Public Overridable ReadOnly Property Operation As String Implements IBatchItem.Operation
    '  Get
    '    Return "No Operation"
    '  End Get
    'End Property

    'Public Property ProcessedBy As String Implements IBatchItem.ProcessedBy
    '  Get
    '    Return mstrProcessedBy
    '  End Get
    '  Set(ByVal value As String)
    '    mstrProcessedBy = value
    '  End Set
    'End Property

    'Public Property ProcessedMessage As String Implements IBatchItem.ProcessedMessage, IOperationExtension.ProcessedMessage
    '  Get
    '    Return mstrProcessedMessage
    '  End Get
    '  Set(ByVal value As String)
    '    Try
    '      mstrProcessedMessage = value
    '      If BatchItem IsNot Nothing Then
    '        BatchItem.ProcessedMessage = value
    '      End If
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Set
    'End Property

    'Public Property ProcessedStatus As ProcessedStatus Implements IBatchItem.ProcessedStatus
    '  Get
    '    Return menumProcessedStatus
    '  End Get
    '  Set(ByVal value As ProcessedStatus)
    '    menumProcessedStatus = value
    '  End Set
    'End Property

    'Public Property ProcessFinishTime As Date Implements IBatchItem.FinishTime
    '  Get
    '    Return mdteProcessFinishTime
    '  End Get
    '  Set(ByVal value As DateTime)
    '    mdteProcessFinishTime = value
    '  End Set
    'End Property

    'Public Property ProcessStartTime As Date Implements IBatchItem.StartTime
    '  Get
    '    Return mdteProcessStartTime
    '  End Get
    '  Set(ByVal value As DateTime)
    '    mdteProcessStartTime = value
    '  End Set
    'End Property

    'Public Property SourceDocId As String Implements IBatchItem.SourceDocId
    '  Get
    '    Return mstrSourceDocId
    '  End Get
    '  Set(ByVal value As String)
    '    mstrSourceDocId = value
    '  End Set
    'End Property

    'Public Property Title As String Implements IBatchItem.Title
    '  Get
    '    Return mstrTitle
    '  End Get
    '  Set(ByVal value As String)
    '    mstrTitle = value
    '  End Set
    'End Property

    'Public Property TotalProcessingTime As String Implements IBatchItem.TotalProcessingTime
    '  Get
    '    Return mstrTotalProcessingTime
    '  End Get
    '  Set(ByVal value As String)
    '    mstrTotalProcessingTime = value
    '  End Set
    'End Property

#End Region

#Region "IOperationInformation Implementation"


    Public Property ExtensionPath As String Implements IOperationInformation.ExtensionPath
      Get
        Return mstrExtensionPath
      End Get
      Set(ByVal value As String)
        mstrExtensionPath = value
      End Set
    End Property

    Public ReadOnly Property IsExtension As Boolean Implements IOperationInformation.IsExtension
      Get
        Return True
      End Get
    End Property

    'Public ReadOnly Property LogSession As Session Implements ILoggable.LogSession
    '  Get
    '  Try
    '    If mobjLogSession Is Nothing Then
    '      InitializeLogSession()
    '    End If
    '    Return mobjLogSession
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    '  End Get
    'End Property

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

    Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
      Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
      Try
        Operation.WriteOperableXml(Me, writer)
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

#Region " IDisposable Support "

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private disposedValue As Boolean     ' To detect redundant calls

    ' IDisposable
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
      Try
        If Not Me.disposedValue Then
          If disposing Then
            MyBase.Dispose(disposing)
            ' DISPOSETODO: free other state (managed objects).

            mstrProcessedMessage = Nothing
            MobjParent = Nothing
            mobjWorkItem = Nothing
            mstrDocumentId = Nothing
            MobjParameters = Nothing
            menuResult = Nothing
            If mobjResultDetail IsNot Nothing Then
              mobjResultDetail.Dispose()
              mobjResultDetail = Nothing
            End If
            mblnLogResult = Nothing
            menuScope = Nothing
            mstrExtensionPath = Nothing
            mdatStartTime = Nothing
            mdatFinishTime = Nothing
            mobjSourceConnection = Nothing
            mobjDestinationConnection = Nothing
            mobjPrimaryConnection = Nothing
            'FinalizeLogSession()
          End If

          ' DISPOSETODO: free your own state (unmanaged objects).
          ' DISPOSETODO: set large fields to null.

        End If
        Me.disposedValue = True
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub MobjParameters_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles MobjParameters.CollectionChanged
      Try
        ' Added by Ernie Bahr on 11/20/2015
        ' Until we can better handle these in the UI components we will filter out any operations with multi-valued parameters

        If e.NewItems IsNot Nothing Then
          If e.NewItems.Count > 0 Then
            ' lpOperation.Parameters.Any(Function(lobjParameter) lobjParameter.Cardinality = Cardinality.ecmMultiValued)
            For Each lobjParameter As IParameter In e.NewItems
              If lobjParameter.Cardinality = Cardinality.ecmMultiValued Then
                'LogSession.LogWarning("Operation '{0}' skipped over.  Operations with multi-valued parameters are currently unsupported.", lobjParameter.Name)
                Throw New InvalidPropertyException("Multi-Valued parameters are currently unsupported.", lobjParameter)
              End If
            Next
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub MobjParameters_ItemPropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles MobjParameters.ItemPropertyChanged
      RaiseEvent ParameterPropertyChanged(sender, e)
    End Sub

    'Public Sub InitializeLogSession() Implements ILoggable.InitializeLogSession
    '  Try
    '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, System.Drawing.Color.Plum)
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

#End Region

  End Class

End Namespace