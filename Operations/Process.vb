' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Process.vb
'  Description :  [type_description_here]
'  Created     :  11/23/2011 4:56:55 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Globalization
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

<TypeConverter(GetType(ExpandableObjectConverter)),
DebuggerDisplay("{DebuggerIdentifier(),nq}")>
Public Class Process
  Implements IProcess
  Implements ISerialize
  Implements IJsonSerializable(Of Process)
  Implements IXmlSerializable
  Implements ICloneable
  Implements IDisposable
  'Implements ILoggable

#Region "Class Constants"

  Private Const PROCESS_FILE_EXTENSION As String = "prf"
  Friend Const LOG_RESULT As String = "LogResult"
  Private Const XML_HEADER As String = "<?xml version=""1.0"" encoding=""utf-16""?>"
  Public Const PARAM_SOURCE_DOC_ID As String = "SourceDocId"

#End Region

#Region "Class Variables"

  Private mstrDescription As String = String.Empty
  Private mstrDisplayName As String = String.Empty
  Private mstrName As String = String.Empty
  Private WithEvents MobjOperations As New Operations
  Private mstrDocumentId As String = String.Empty
  Private mobjParameters As IParameters = New Parameters
  Private mstrOriginalFilePath As String = String.Empty
  Private mobjParent As IItemParent = Nothing
  Private mobjWorkItem As IWorkItem = Nothing
  Private mstrProcessedMessage As String = String.Empty
  Private ReadOnly mobjFailureOperations As New Operations
  Private menuResult As Result = OperationEnumerations.Result.NotProcessed
  Private mblnLogResult As Boolean = True
  Private mdatStartTime As DateTime = DateTime.MinValue
  Private mdatFinishTime As DateTime = DateTime.MinValue
  Private ReadOnly mobjLocale As CultureInfo = CultureInfo.CurrentCulture
  Private mstrLocale As String = mobjLocale.Name
  Private mobjResultDetails As IProcessResult = Nothing
  Private mobjHost As Object = Nothing
  Private mobjTag As Object = Nothing
  Private mblnIsEmpty As Boolean = True

  Private WithEvents MobjCurrentOperation As IOperable = Nothing
  Private mobjRunBeforeBegin As IOperable = Nothing
  Private mobjRunAfterComplete As IOperable = Nothing
  Private mobjRunOnFailure As IOperable = Nothing

  Private mobjRunBeforeParentBegin As IOperable = Nothing
  Private mobjRunAfterParentComplete As IOperable = Nothing
  Private mobjRunOnParentFailure As IOperable = Nothing

  Private WithEvents MobjRunBeforeJobBegin As IOperable = Nothing
  Private mobjRunAfterJobComplete As IOperable = Nothing
  Private mobjRunOnJobFailure As IOperable = Nothing

  Private mstrInstanceId As String = String.Empty
  Private mblnAutoRollback As Boolean = False

  ' Private mstrParameterExpression As String = "(?<Prefix>.*){Param:(?<Param>[a-zA-Z0-9]*)}(?<Suffix>.*)"
  Private Shared ReadOnly mstrParameterExpression As String = "(?<Prefix>.*){(?<ParamName>[a-zA-Z0-9]*):(?<ParamValue>[a-zA-Z0-9]*)}(?<Suffix>.*)"

#End Region

#Region "Constructors"

  Public Sub New()
    Try
      'InitializeLogSession()
      mstrInstanceId = GenerateInstanceId()
      mblnIsEmpty = True
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '   Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(ByVal lpFilePath As String)

    Dim lstrErrorMessage As String = String.Empty
    Dim lobjProcess As IProcess = Nothing

    Try
      'InitializeLogSession()
      lobjProcess = CType(Deserialize(lpFilePath, lstrErrorMessage), IProcess)

      If lobjProcess Is Nothing Then
        ' Check the error message
        If lstrErrorMessage.Length > 0 Then
          Throw New Exception(lstrErrorMessage)
        End If
      End If

      Helper.AssignObjectProperties(lobjProcess, Me)

      ' Additional Assignments
      With lobjProcess
        mstrName = .Name
        mstrDisplayName = .DisplayName
        mobjParameters = .Parameters
        MobjOperations = CType(.Operations, Operations)

        mobjRunBeforeBegin = .RunBeforeBegin
        MobjRunBeforeJobBegin = .RunBeforeJobBegin
        mobjRunBeforeParentBegin = .RunBeforeParentBegin

        mobjRunAfterComplete = .RunAfterComplete
        mobjRunAfterJobComplete = .RunAfterJobComplete
        mobjRunAfterParentComplete = .RunAfterParentComplete

        mobjRunOnFailure = .RunOnFailure
        mobjRunOnJobFailure = .RunOnJobFailure
        mobjRunOnParentFailure = .RunOnParentFailure

      End With

      mstrInstanceId = GenerateInstanceId()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Friend Sub New(ByVal lpName As String, ByVal lpDescription As String)
    Try
      'InitializeLogSession()
      mstrInstanceId = GenerateInstanceId()
      mstrName = lpName
      mstrDescription = lpDescription
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Friend Sub New(ByVal lpName As String, ByVal lpDescription As String, lpLocale As CultureInfo)
    Try
      'InitializeLogSession()
      mstrInstanceId = GenerateInstanceId()
      mstrName = lpName
      mstrDescription = lpDescription
      mobjLocale = lpLocale
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Friend Sub New(ByVal lpName As String, ByVal lpDisplayName As String, ByVal lpDescription As String)
    Try
      'InitializeLogSession()
      mstrInstanceId = GenerateInstanceId()
      mstrName = lpName
      mstrDisplayName = lpDisplayName
      mstrDescription = lpDescription
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  ''' <summary>
  ''' Creates a new process by cloning the incoming process.
  ''' </summary>
  ''' <param name="lpProcess"></param>
  ''' <remarks>
  ''' This is used everytime a process is updated to propagate 
  ''' clones of the process to each batch.  It is very important 
  ''' to update this method each time a new property is added to 
  ''' the process class, otherwise the process clones in the 
  ''' batches will not have the values set on the new properties.
  ''' </remarks>
  Friend Sub New(ByVal lpProcess As IProcess)
    Try
      'InitializeLogSession()
      mstrInstanceId = GenerateInstanceId()
      mstrName = CStr(lpProcess.Name.Clone)
      mstrDisplayName = CStr(lpProcess.DisplayName.Clone)
      mstrDescription = CStr(lpProcess.Description.Clone)
      MobjOperations = CType(lpProcess.Operations.Clone, Operations)

      ' Clone the events
      ' Clone the direct process events
      If lpProcess.RunBeforeBegin IsNot Nothing Then
        mobjRunBeforeBegin = CType(lpProcess.RunBeforeBegin.Clone, IOperable)
      End If
      If lpProcess.RunAfterComplete IsNot Nothing Then
        mobjRunAfterComplete = CType(lpProcess.RunAfterComplete.Clone, IOperable)
      End If
      If lpProcess.RunOnFailure IsNot Nothing Then
        mobjRunOnFailure = CType(lpProcess.RunOnFailure.Clone, IOperable)
      End If

      ' Clone the parent events
      If lpProcess.RunBeforeParentBegin IsNot Nothing Then
        mobjRunBeforeParentBegin = CType(lpProcess.RunBeforeParentBegin.Clone, IOperable)
      End If
      If lpProcess.RunAfterParentComplete IsNot Nothing Then
        mobjRunAfterParentComplete = CType(lpProcess.RunAfterParentComplete.Clone, IOperable)
      End If
      If lpProcess.RunOnParentFailure IsNot Nothing Then
        mobjRunOnParentFailure = CType(lpProcess.RunOnParentFailure.Clone, IOperable)
      End If

      ' Clone the job events
      If lpProcess.RunBeforeJobBegin IsNot Nothing Then
        MobjRunBeforeJobBegin = CType(lpProcess.RunBeforeJobBegin.Clone, IOperable)
      End If
      If lpProcess.RunAfterJobComplete IsNot Nothing Then
        mobjRunAfterJobComplete = CType(lpProcess.RunAfterJobComplete.Clone, IOperable)
      End If
      If lpProcess.RunOnJobFailure IsNot Nothing Then
        mobjRunOnJobFailure = CType(lpProcess.RunOnJobFailure.Clone, IOperable)
      End If

      ' <Added by: Ernie at: 8/13/2012-2:10:12 PM on machine: ERNIE-M4400>
      ' For the Host and tag we will not clone the values since 
      ' they are references that should be persisted.  If it turns 
      ' out that we need to clone the tag then we can change it 
      ' but the host should not get cloned.
      If lpProcess.Host IsNot Nothing Then
        mobjHost = lpProcess.Host
      End If

      If lpProcess.Tag IsNot Nothing Then
        mobjTag = lpProcess.Tag
      End If
      ' </Added by: Ernie at: 8/13/2012-2:10:12 PM on machine: ERNIE-M4400>

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Properties"

  'Public ReadOnly Property CanRollback As Boolean
  '	Get
  '		Try
  '			For Each lobjOperable As IOperable In Me.Operations
  '				If lobjOperable.CanRollback = True Then
  '					Return True
  '				End If
  '			Next
  '			Return False
  '		Catch ex As Exception
  '			ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '			'   Re-throw the exception to the caller
  '			Throw
  '		End Try
  '	End Get
  'End Property

  Public ReadOnly Property IsEmpty As Boolean Implements IProcess.IsEmpty
    Get
      Return mblnIsEmpty
    End Get
  End Property

  Public ReadOnly Property IsDisposed() As Boolean Implements IProcess.IsDisposed
    Get
      Return disposedValue
    End Get
  End Property

  Public ReadOnly Property OriginalFilePath() As String
    Get
      Return mstrOriginalFilePath
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
        Operations.AssignHost(value)
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
        Operations.AssignTag(value)
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

  <Category(Operation.CATEGORY_BEHAVIOR)>
  Public Property RunAfterJobComplete As IOperable Implements IProcess.RunAfterJobComplete
    Get
      Try
        Return mobjRunAfterJobComplete
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IOperable)
      Try
        mobjRunAfterJobComplete = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  <Category(Operation.CATEGORY_BEHAVIOR)>
  Public Property RunAfterParentComplete As IOperable Implements IProcess.RunAfterParentComplete
    Get
      Try
        Return mobjRunAfterParentComplete
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IOperable)
      Try
        mobjRunAfterParentComplete = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  <Category(Operation.CATEGORY_BEHAVIOR)>
  Public Property RunBeforeJobBegin As IOperable Implements IProcess.RunBeforeJobBegin
    Get
      Try
        Return MobjRunBeforeJobBegin
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IOperable)
      Try
        MobjRunBeforeJobBegin = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  <Category(Operation.CATEGORY_BEHAVIOR)>
  Public Property RunBeforeParentBegin As IOperable Implements IProcess.RunBeforeParentBegin
    Get
      Try
        Return mobjRunBeforeParentBegin
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IOperable)
      Try
        mobjRunBeforeParentBegin = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  <Category(Operation.CATEGORY_BEHAVIOR)>
  Public Property RunOnJobFailure As IOperable Implements IProcess.RunOnJobFailure
    Get
      Try
        Return mobjRunOnJobFailure
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IOperable)
      Try
        mobjRunOnJobFailure = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  <Category("Behavior")>
  Public Property RunOnParentFailure As IOperable Implements IProcess.RunOnParentFailure
    Get
      Try
        Return mobjRunOnParentFailure
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IOperable)
      Try
        mobjRunOnParentFailure = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public ReadOnly Property InstanceId As String Implements IOperable.InstanceId
    Get
      Return mstrInstanceId
    End Get
  End Property

  ''' <summary>
  '''     Indicates whether or not the process should attempt to automatically rollback on failure.
  ''' </summary>
  ''' <value>
  '''     <para>
  '''         
  '''     </para>
  ''' </value>
  ''' <remarks>
  '''     
  ''' </remarks>
  Public Property AutoRollback As Boolean Implements IProcess.AutoRollback
    Get
      Try
        Return mblnAutoRollback
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Boolean)
      Try
        mblnAutoRollback = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public ReadOnly Property CanRollback As Boolean Implements IOperable.CanRollback
    Get
      Try
        ' If at least one of the operations can rollback 
        ' then the process can at least do a partial rollback.
        For Each lobjOperable As IOperable In Operations
          If lobjOperable.CanRollback Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

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

  Public Shared Function GetInlineParameter(ByVal lpValue As String) As INameValuePair
    Try

      Dim lobjRegex As New Regex(mstrParameterExpression,
          RegexOptions.CultureInvariant Or RegexOptions.Compiled)

      ' Split the InputText wherever the regex matches
      Dim lstrResults As String() = lobjRegex.Split(lpValue)

      ' Test to see if there is a match in the InputText
      Dim lblnIsMatch As Boolean = lobjRegex.IsMatch(lpValue)

      If lblnIsMatch Then
        Dim lintParameterNameGroupNumber As Integer = lobjRegex.GroupNumberFromName("ParamName")
        Dim lintParameterValueGroupNumber As Integer = lobjRegex.GroupNumberFromName("ParamValue")

        If lintParameterNameGroupNumber > 0 AndAlso lintParameterValueGroupNumber > 0 Then
          Return New Configuration.KeyValuePair(lstrResults(lintParameterNameGroupNumber),
                                                lstrResults(lintParameterValueGroupNumber))
        Else
          Return Nothing
        End If

      Else
        Return Nothing
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  'Public Function ResolveInlineParameter(ByVal lpValue As String) As String Implements IOperable.ResolveInlineParameter
  '  Try
  '    Return Operation.ResolveInlineParameter(Me, lpValue)
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  Public Overloads Function ToString() As String Implements IOperable.ToString
    Try
      Return DebuggerIdentifier()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function FromXml(lpProcessXml As String) As Process

    Dim lstrErrorMessage As String = String.Empty
    Dim lobjProcess As IProcess = Nothing
    Try

      lobjProcess = Serializer.Deserialize.XmlString(lpProcessXml, GetType(Process))

      If lobjProcess Is Nothing Then
        ' Check the error message
        If lstrErrorMessage.Length > 0 Then
          Throw New Exception(lstrErrorMessage)
        End If
      End If

      Return lobjProcess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Protected Methods"

  Protected Sub UpdateParameterToEnum(lpName As String, lpEnumType As Type) Implements IOperable.UpdateParameterToEnum
    Try
      ParameterFactory.UpdateParameterToEnum(Me.Parameters, lpName, lpEnumType)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Friend Overridable Function DebuggerIdentifier() As String
    Dim lobjIdentifierBuilder As New Text.StringBuilder
    Try

      If disposedValue = True Then
        Return "Process Disposed"
      End If

      If Not String.IsNullOrEmpty(Name) Then
        lobjIdentifierBuilder.AppendFormat("{0}", Name)
      End If

      If Operations.Count = 0 OrElse Operations Is Nothing Then
        If String.IsNullOrEmpty(Name) Then
          lobjIdentifierBuilder.Append("Empty Process: No Operations")
        Else
          lobjIdentifierBuilder.Append(": No Operations")
        End If
      ElseIf Operations.Count = 1 Then
        lobjIdentifierBuilder.Append(": 1 Operation")
      Else
        lobjIdentifierBuilder.AppendFormat(": {0} Operations", Operations.Count)
      End If

      Return lobjIdentifierBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Return lobjIdentifierBuilder.ToString
    End Try
  End Function

#End Region

#Region "PrivateMethods"

  Private Shared Function GenerateInstanceId() As String
    Try
      Return Guid.NewGuid().ToString
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Sub Process_OperatingError(ByVal sender As Object, ByVal e As OperableErrorEventArgs) Handles Me.OperatingError
    Try
      Me.WorkItem.ProcessedMessage = e.WorkItem.ProcessedMessage
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub MobjOperations_CollectionChanged(sender As Object, e As Specialized.NotifyCollectionChangedEventArgs) Handles MobjOperations.CollectionChanged
    Try
      If MobjOperations.Count = 0 Then
        mblnIsEmpty = True
      Else
        mblnIsEmpty = False
      End If
      If e.Action = Specialized.NotifyCollectionChangedAction.Add Then
        For Each lobjOperation As IOperation In e.NewItems
          lobjOperation.OperableParent = Me
        Next
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '   Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "IProcess Implementation"

  Public Event Complete(ByVal sender As Object, ByVal e As OperableEventArgs) Implements IOperable.Complete

  Public Event Begin(ByVal sender As Object, ByVal e As OperableEventArgs) Implements IOperable.Begin

  Public Event OperatingError(ByVal sender As Object, ByVal e As OperableErrorEventArgs) Implements IOperable.OperatingError
  Public Event ParameterPropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements IOperable.ParameterPropertyChanged

  Public Overridable Sub CheckParameters() Implements IOperable.CheckParameters
    Try
      ' Do Nothing
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  <Category("ReadOnly Properties")>
  Public ReadOnly Property Description As String Implements IOperable.Description
    Get
      Return mstrDescription
    End Get
    'Set(value As String)
    '  mstrDescription = value
    'End Set
  End Property

  <Category("ReadOnly Properties")>
  Public ReadOnly Property DisplayName As String Implements IOperable.DisplayName ', IProcess.DisplayName
    Get
      Try
        If String.IsNullOrEmpty(mstrDisplayName) Then
          mstrDisplayName = Helper.CreateDisplayName(Me.Name)
        End If
        Return mstrDisplayName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    'Set(value As String)
    '  mstrDisplayName = value
    'End Set
  End Property

  <Category("ReadOnly Properties")>
  Public ReadOnly Property Name As String Implements IProcess.Name
    Get
      Return mstrName
    End Get
    'Set(value As String)
    '  mstrName = value
    'End Set
  End Property

  <Category("Behavior"), DisplayName("Operations"),
  Description("The collection of operations for the process.")>
  Public ReadOnly Property Operations As IOperations Implements IProcess.Operations
    Get
      Return MobjOperations
    End Get
  End Property

  <Category("ReadOnly Properties")>
  Public ReadOnly Property ResultDetail As IProcessResult Implements IProcess.ResultDetail
    Get
      Try
        Return mobjResultDetails
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  <Category("ReadOnly Properties")>
  Public ReadOnly Property OperableResultDetail As IOperableResult Implements IOperable.ResultDetail
    Get
      Try
        Return mobjResultDetails
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  <Category("ReadWrite Properties")>
  Public Property DocumentId As String Implements IOperable.DocumentId
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
        Return Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As IOperable)
      Try
        '  Do nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  <Category("ReadWrite Properties")>
  Public Property Parent As IItemParent Implements IProcess.Parent
    Get
      Return mobjParent
    End Get
    Set(ByVal value As IItemParent)
      mobjParent = value
    End Set
  End Property

  <Category("ReadWrite Properties")>
  Public Property WorkItem As IWorkItem Implements IProcess.WorkItem
    Get
      Return mobjWorkItem
    End Get
    Set(ByVal value As IWorkItem)
      mobjWorkItem = value
    End Set
  End Property

  ''' <summary>
  ''' Executes the process
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Execute(ByVal lpWorkItem As IWorkItem) As OperationEnumerations.Result Implements IOperable.Execute
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

      OnComplete(New OperableEventArgs(Me, lpWorkItem))

      ' If applicable, run any post-operation(s)
      If RunAfterComplete IsNot Nothing Then
        lpWorkItem.ProcessedStatus = menuResult
        Dim lenuRunAfterCompleteResult As Result = RunAfterComplete.Execute(lpWorkItem)
        If lenuRunAfterCompleteResult = OperationEnumerations.Result.Failed Then
          OnError(New OperableErrorEventArgs(Me, lpWorkItem, lpWorkItem.ProcessedMessage))
        End If
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      OnError(New OperableErrorEventArgs(Me, lpWorkItem, ex))
    Finally
      'LogSession.LeaveMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
    End Try

    Return menuResult

  End Function

  'Public Function Rollback(lpWorkItem As IWorkItem) As Result Implements IProcess.Rollback
  '	Try
  '		Dim lobjResult As New Result

  '		' Walk backwards through the operations and try to rollback each one
  '		For lintOperationCounter As Integer = Operations.Count - 1 To 0 Step -1
  '			If Operations(lintOperationCounter).CanRollback = True Then
  '				Operations(lintOperationCounter).Rollback(lpWorkItem)
  '			End If
  '		Next
  '	Catch ex As Exception
  '		ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '		'   Re-throw the exception to the caller
  '		Throw
  '	End Try
  'End Function

  Public Function Rollback(ByVal lpWorkItem As IWorkItem) As Result Implements IOperable.Rollback
    Try
      Return OnRollback()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Sub SetDescription(ByVal lpDescription As String) Implements IOperable.SetDescription ', IActionItem.SetDescription
    Try
      mstrDescription = lpDescription
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub SetResult(ByVal lpResult As Result) Implements IProcess.SetResult
    Try
      menuResult = lpResult
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overridable Sub Reset() Implements IProcess.Reset
    Try
      menuResult = Result.NotProcessed
      mobjResultDetails?.Dispose()
      mobjResultDetails = Nothing
      mdatStartTime = DateTime.MinValue
      mdatFinishTime = DateTime.MinValue
      For Each lobjOperableStep As IOperable In Operations
        lobjOperableStep.Reset()
      Next
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Overridable Function OnExecute() As OperationEnumerations.Result Implements IProcess.OnExecute
    Try

      Dim lenuStepResult As OperationEnumerations.Result = OperationEnumerations.Result.NotProcessed
      Me.ProcessedMessage = String.Empty

      If Operations.Count = 0 Then
        lenuStepResult = OperationEnumerations.Result.Failed
        Me.ProcessedMessage = "No operations were found to execute in this process."
        Return lenuStepResult
      End If

      ' Syncronize the host and tag to all the child operations
      Operations.AssignHost(Me.Host)
      Operations.AssignTag(Me.Tag)

      For Each lobjOperableStep As IOperable In Operations
        MobjCurrentOperation = lobjOperableStep
        lobjOperableStep.ProcessedMessage = String.Empty
        'LogSession.LogDebug("About to execute process operation '{0}'.", lobjOperableStep.Name)
        lenuStepResult = lobjOperableStep.Execute(Me.WorkItem)
        If lenuStepResult = OperationEnumerations.Result.Failed Then
          If Not String.IsNullOrEmpty(lobjOperableStep.ProcessedMessage) Then
            Me.ProcessedMessage = lobjOperableStep.ProcessedMessage
            'LogSession.LogDebug("Process operation '{0}' failed ({1}).", lobjOperableStep.Name, lobjOperableStep.ProcessedMessage)
          Else
            'LogSession.LogDebug("Process operation '{0}' failed with no message.", lobjOperableStep.Name)
          End If
          ' Execute the RunOnFailure operation(s)
          lobjOperableStep.RunOnFailure?.Execute(Me.WorkItem)
          Exit For
        End If
      Next

      If Not String.IsNullOrEmpty(Me.WorkItem.ProcessedMessage) Then
        Me.ProcessedMessage = Me.WorkItem.ProcessedMessage
      End If

      menuResult = lenuStepResult

      Return menuResult

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Overridable Function OnRollback() As Result Implements IProcess.OnRollBack
    Try
      Dim lenuResult As Result
      For Each lobjOperable As IOperable In Operations
        If lobjOperable.CanRollback Then
          lenuResult = lobjOperable.Rollback(Me.WorkItem)
        End If
      Next
      Return lenuResult
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  'Protected Overridable Function OnRollback() As Result Implements IProcess.OnRollback
  '	Try

  '	Catch ex As Exception
  '		ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '		'   Re-throw the exception to the caller
  '		Throw
  '	End Try
  'End Function

  <Category(Operation.CATEGORY_CONFIG)>
  Public Property Parameters As IParameters Implements IOperable.Parameters
    Get
      Return mobjParameters
    End Get
    Set(ByVal value As IParameters)
      mobjParameters = value
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

  <Category("ReadWrite Properties")>
  Public Property ProcessedMessage As String Implements IProcess.ProcessedMessage
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

  <Category("ReadOnly Properties")>
  Public ReadOnly Property Result As OperationEnumerations.Result Implements IOperable.Result
    Get
      Return menuResult
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
  <Category("ReadWrite Properties")>
  Public Property LogResult As Boolean Implements IOperable.LogResult
    Get
      Return mblnLogResult
    End Get
    Set(ByVal value As Boolean)
      Try
        mblnLogResult = value
        ' Syncronize the value to each child operation as well.
        For Each lobjOperation As IOperable In Me.Operations
          lobjOperation.LogResult = value
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  ''' <summary>
  '''     This should be considered a read-only property for Process.
  ''' </summary>
  ''' <value>
  '''     <para>
  '''         Will always return True.
  '''     </para>
  ''' </value>
  ''' <remarks>
  '''     
  ''' </remarks>
  <Category("ReadWrite Properties")>
  Public Property ShouldExecute As Boolean Implements IOperable.ShouldExecute
    Get
      Try
        Return True
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As Boolean)
      Try
        ' Ignore
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  ''' <summary>
  ''' Gets or sets the time when the process is started.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  <Category("ReadWrite Properties")>
  Public Property StartTime As DateTime Implements IProcess.StartTime
    Get
      Return mdatStartTime
    End Get
    Set(ByVal value As DateTime)
      Try

        mdatStartTime = value

        ' If not already set, syncronize the Process time with the WorkItem time.
        If Me.WorkItem IsNot Nothing AndAlso Me.WorkItem.StartTime = DateTime.MinValue Then
          Me.WorkItem.StartTime = value
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  ''' <summary>
  ''' Gets or sets the time when the process is finished.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  <Category("ReadWrite Properties")>
  Public Property FinishTime As DateTime Implements IProcess.FinishTime
    Get
      Return mdatFinishTime
    End Get
    Set(ByVal value As DateTime)
      Try

        mdatFinishTime = value

        ' If not already set, syncronize the Process time with the WorkItem time.
        If Me.WorkItem IsNot Nothing AndAlso Me.WorkItem.FinishTime = DateTime.MinValue Then
          Me.WorkItem.FinishTime = value
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  ''' <summary>
  ''' Gets the total processing time for the process.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  <Category("ReadOnly Properties")>
  Public ReadOnly Property TotalProcessingTime As TimeSpan Implements IProcess.TotalProcessingTime
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

  Public Overridable Sub OnBegin(ByVal e As OperableEventArgs) Implements IOperable.OnBegin
    Try

      Reset()

      ' Since all the Operation subclasses should call this method before executing
      ' we can ensure the initialization of the WorkItem and Parent properties here.
      WorkItem = e.WorkItem

      ' Set the start time
      StartTime = Now

      If Parent Is Nothing AndAlso WorkItem IsNot Nothing AndAlso WorkItem.Parent IsNot Nothing Then
        Parent = WorkItem.Parent
      End If

      For Each lobjOperation As IOperable In Operations
        lobjOperation.SetInstanceId(Me.InstanceId)
      Next

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
      mobjResultDetails = New ProcessResult(Me)

      If WorkItem.ProcessedStatus = ProcessedStatus.NotProcessed Then
        Select Case Result
          Case Result.Failed
            WorkItem.ProcessedStatus = ProcessedStatus.Failed
          Case Result.Success
            WorkItem.ProcessedStatus = ProcessedStatus.Success
        End Select
      End If


      ' Make sure we dispose of the document attachment to release any associated memory
      If Me.WorkItem IsNot Nothing AndAlso Me.WorkItem.Document IsNot Nothing AndAlso Me.WorkItem.Document.IsDisposed = False Then
        Me.WorkItem.Document.Dispose()
        Me.WorkItem.Document = Nothing
      End If

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
        Me.ProcessedMessage = e.Exception.Message
      End If

      ' Initialize the detailed results
      mobjResultDetails = New ProcessResult(Me)

      ' Make sure we dispose of the document attachment to release any associated memory
      If Me.WorkItem IsNot Nothing AndAlso Me.WorkItem.Document IsNot Nothing Then
        If AutoRollback = True Then
          Rollback(Me.WorkItem)
        End If
        Me.WorkItem.Document.Dispose()
      End If

      If Not String.IsNullOrEmpty(Me.ProcessedMessage) Then
        'LogSession.LogError(Me.ProcessedMessage)
      End If

      RaiseEvent OperatingError(Me, e)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "ISerialize Implementation"

  Public ReadOnly Property DefaultFileExtension As String Implements ISerialize.DefaultFileExtension
    Get
      Try
        Return PROCESS_FILE_EXTENSION
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
    Try

#If NET8_0_OR_GREATER Then
      ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

      mstrOriginalFilePath = lpFilePath
      Dim lobjProcess As Process = CType(Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType), Process)

      If DisplayName.Length = 0 AndAlso OriginalFilePath.Length > 0 Then
        lobjProcess.mstrDisplayName = Helper.CreateDisplayName(IO.Path.GetFileNameWithoutExtension(OriginalFilePath))
      End If

      Return lobjProcess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function Deserialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
    Try

#If NET8_0_OR_GREATER Then
      ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

      Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Shared Function Deserialize(ByVal lpXmlNode As XmlNode) As IProcess
    Try
      Return CType(Serializer.Deserialize.XmlString(lpXmlNode.OuterXml, GetType(Process)), IProcess)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
    Try
      Return Helper.FormatXmlDocument(Serializer.Serialize.Xml(Me))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
    Try

      If lpFileExtension.Length = 0 Then
        ' No override was provided
        If lpFilePath.EndsWith(DefaultFileExtension) = False Then
          lpFilePath = lpFilePath.Remove(lpFilePath.Length - 3) & DefaultFileExtension
        End If

      End If

      ' If no display name was set previously, set it to the requested file name
      If DisplayName.Length = 0 Then
        mstrDisplayName = Helper.CreateDisplayName(IO.Path.GetFileNameWithoutExtension(lpFilePath))
      End If

      Serializer.Serialize.XmlFile(Me, lpFilePath)

      Helper.FormatXmlFile(lpFilePath)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub Save(ByVal lpFilePath As String)
    Try
      Serialize(lpFilePath, IO.Path.GetExtension(lpFilePath))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
    Try
      Serialize(lpFilePath, String.Empty)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize
    Try
      If lpWriteProcessingInstruction = True Then
        Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
      Else
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      End If

      Helper.FormatXmlFile(lpFilePath)

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

  Public Function ToXmlString() As String Implements ISerialize.ToXmlString, IProcess.ToXmlString
    Try
      Return Helper.FormatXmlString(Serializer.Serialize.XmlString(Me))
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

  Public Function ToXmlElementString() As String Implements IOperable.ToXmlElementString
    Try
      Return Serializer.Serialize.XmlElementString(Me)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "IJsonSerializable(Of Process)"

  Public Overloads Function ToJson() As String Implements IJsonSerializable(Of Process).ToJson, IOperable.ToJson
    Try
      Return JsonConvert.SerializeObject(Me, Newtonsoft.Json.Formatting.None, New ProcessJsonConverter())
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function FromJson(lpJson As String) As Process Implements IJsonSerializable(Of Process).FromJson
    Try
      Return JsonConvert.DeserializeObject(lpJson, GetType(Process), New ProcessJsonConverter())
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

  Public Shared Function CreateFromJson(lpJson As String) As Process
    Try
      Return JsonConvert.DeserializeObject(lpJson, GetType(Process), New ProcessJsonConverter())
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#Region "IXmlSerializable Implementation"



  Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
    ' As per the Microsoft guidelines this is not implemented
    Return Nothing
  End Function

  Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml

    Dim lobjProcessXmlBuilder As New StringBuilder
    Dim lobjXmlDocument As New XmlDocument
    Dim lobjAttribute As XmlAttribute = Nothing
    Dim lstrOperationType As String = Nothing
    Dim lobjOperable As IOperable = Nothing
    Dim lobjParameterNodes As XmlNodeList = Nothing
    Dim lobjParameter As Parameter = Nothing

    Try

      ' <Modified by: Ernie Bahr at 11/13/2012-07:50:54 on machine: ERNIEBAHR-THINK>
      ' We were having problems reading when loading as part of a larger object such as JobConfiguration.  
      ' Recreating the xml string seems to resolve the issue.
      lobjProcessXmlBuilder.AppendLine(XML_HEADER)
      lobjProcessXmlBuilder.Append(reader.ReadOuterXml)

      'lobjXmlDocument.Load(reader)
      lobjXmlDocument.LoadXml(lobjProcessXmlBuilder.ToString)
      ' </Modified by: Ernie Bahr at 11/13/2012-07:50:54 on machine: ERNIEBAHR-THINK>

      With lobjXmlDocument
        ' Get the name
        mstrName = .DocumentElement.GetAttribute("Name")

        ' Get the description
        mstrDescription = .DocumentElement.GetAttribute("Description")

        ' Get the LogResult property
        Boolean.TryParse(.DocumentElement.GetAttribute(LOG_RESULT), Me.LogResult)

        ' Try to get the times
        ' Get the locale
        mstrLocale = .DocumentElement.GetAttribute("locale")

        ' Get the StartTime
        Me.StartTime = Helper.FromDetailedDateString(.DocumentElement.GetAttribute("StartTime"), mstrLocale)

        ' Get the FinishTime
        Me.FinishTime = Helper.FromDetailedDateString(.DocumentElement.GetAttribute("FinishTime"), mstrLocale)

        ' Read the Operation elements
        Dim lobjOperationsNode As XmlNode = .SelectSingleNode("//Operations")

        Operations.AddRange(GetOperations(lobjOperationsNode))

        ' Read the event elements
        Dim lobjRunBeforeBeginNode As XmlNode = .DocumentElement.SelectSingleNode("RunBeforeBegin")
        If lobjRunBeforeBeginNode IsNot Nothing Then
          RunBeforeBegin = OperationFactory.GetRunOperable(lobjRunBeforeBeginNode)
        End If

        Dim lobjRunAfterCompleteNode As XmlNode = .DocumentElement.SelectSingleNode("RunAfterComplete")
        If lobjRunAfterCompleteNode IsNot Nothing Then
          RunAfterComplete = OperationFactory.GetRunOperable(lobjRunAfterCompleteNode)
        End If

        Dim lobjRunOnFailureNode As XmlNode = .DocumentElement.SelectSingleNode("RunOnFailure")
        If lobjRunOnFailureNode IsNot Nothing Then
          RunOnFailure = OperationFactory.GetRunOperable(lobjRunOnFailureNode)
        End If

        Dim lobjRunBeforeParentBeginNode As XmlNode = .DocumentElement.SelectSingleNode("RunBeforeParentBegin")
        If lobjRunBeforeParentBeginNode IsNot Nothing Then
          RunBeforeParentBegin = OperationFactory.GetRunOperable(lobjRunBeforeParentBeginNode)
        End If

        Dim lobjRunAfterParentCompleteNode As XmlNode = .DocumentElement.SelectSingleNode("RunAfterParentComplete")
        If lobjRunAfterParentCompleteNode IsNot Nothing Then
          RunAfterParentComplete = OperationFactory.GetRunOperable(lobjRunAfterParentCompleteNode)
        End If

        Dim lobjRunOnParentFailureNode As XmlNode = .DocumentElement.SelectSingleNode("RunOnParentFailure")
        If lobjRunOnParentFailureNode IsNot Nothing Then
          RunOnParentFailure = OperationFactory.GetRunOperable(lobjRunOnParentFailureNode)
        End If

        Dim lobjRunBeforeJobBeginNode As XmlNode = .DocumentElement.SelectSingleNode("RunBeforeJobBegin")
        If lobjRunBeforeJobBeginNode IsNot Nothing Then
          RunBeforeJobBegin = OperationFactory.GetRunOperable(lobjRunBeforeJobBeginNode)
        End If

        Dim lobjRunAfterJobCompleteNode As XmlNode = .DocumentElement.SelectSingleNode("RunAfterJobComplete")
        If lobjRunAfterJobCompleteNode IsNot Nothing Then
          RunAfterJobComplete = OperationFactory.GetRunOperable(lobjRunAfterJobCompleteNode)
        End If

        Dim lobjRunOnJobFailureNode As XmlNode = .DocumentElement.SelectSingleNode("RunOnJobFailure")
        If lobjRunOnJobFailureNode IsNot Nothing Then
          RunOnJobFailure = OperationFactory.GetRunOperable(lobjRunOnJobFailureNode)
        End If

      End With

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    Finally
      lobjProcessXmlBuilder = Nothing
      lobjXmlDocument = Nothing
    End Try
  End Sub

  Public Shared Function GetOperations(ByVal lpOperationsNode As XmlNode) As IOperations
    Try

      Dim lobjOperations As IOperations = New Operations
      Dim lobjOperable As IOperable = Nothing

      For Each lobjOperation As XmlNode In lpOperationsNode.ChildNodes

        '' We will try to call the serialization of the operable object
        'lobjOperable = Serializer.Deserialize.SoapString( _
        '  .SelectSingleNode("//ProviderSystem").OuterXml, GetType(ProviderSystem))

        ' Read the operation type
        ' This assumes that the operation type is the first attribute
        If lobjOperation.Attributes IsNot Nothing AndAlso lobjOperation.Attributes.Count > 0 Then
          'lstrOperationType = lobjOperation.Attributes(0).Value
          'lstrOperationType = lobjOperation.Name
        Else
          Throw New InvalidOperationException("The operation xml has no type attribute.")
        End If

        lobjOperable = OperationFactory.Create(lobjOperation, CultureInfo.CurrentCulture.Name)

        lobjOperations.Add(lobjOperable)

      Next

      Return lobjOperations

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Shared Function GetOperable(ByVal lpOperationsNode As XmlNode) As IOperable
    Try

      Dim lobjOperable As IOperable = Nothing

      lobjOperable = OperationFactory.Create(lpOperationsNode, CultureInfo.CurrentCulture.Name)

      Return lobjOperable

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
    Try

      'Dim lstrFileName As String

      'If (TypeOf (writer) Is EnhancedXmlTextWriter = False) AndAlso (Helper.CallStackContainsMethodName("ToStream")) Then
      '  lstrFileName = String.Format("{0}.{1}", Me.Name, PROCESS_FILE_EXTENSION)
      'Else
      '  lstrFileName = CType(writer, EnhancedXmlTextWriter).FileName
      'End If

      With writer

        '.WriteAttributeString("xmlns", "xsi", "http://www.w3.org/2001/XMLSchema-instance")
        '.WriteAttributeString("xmlns", "xsd", "http://www.w3.org/2001/XMLSchema")

        ' Write the Process Name attribute
        .WriteAttributeString("Name", Me.Name)

        ' Write the Process Description attribute
        .WriteAttributeString("Description", Me.Description)

        ' Write the LogResult attribute
        .WriteAttributeString(LOG_RESULT, Me.LogResult.ToString)

        ' Write the result, if applicable
        If Me.Result <> OperationEnumerations.Result.NotProcessed AndAlso Me.LogResult = True Then
          .WriteAttributeString("Result", Me.Result.ToString)
        End If

        ' Write the result, if applicable
        If Me.Result <> OperationEnumerations.Result.NotProcessed AndAlso Me.LogResult = True Then
          .WriteAttributeString("ProcessedMessage", Me.ProcessedMessage)
        End If

        ' Write the times, if applicable
        If Me.StartTime <> DateTime.MinValue AndAlso Me.LogResult = True Then
          .WriteAttributeString("StartTime", Helper.ToDetailedDateString(Me.StartTime, mstrLocale))
        End If

        If Me.FinishTime <> DateTime.MinValue AndAlso Me.LogResult = True Then
          .WriteAttributeString("FinishTime", Helper.ToDetailedDateString(Me.FinishTime, mstrLocale))
        End If

        If Me.TotalProcessingTime <> TimeSpan.Zero AndAlso Me.LogResult = True Then
          .WriteAttributeString("TotalProcessingTime", Me.TotalProcessingTime.ToString)
        End If

        If Me.LogResult Then
          .WriteAttributeString("locale", Me.Locale.Name)
        End If

        ' Open the Parameters Element
        .WriteStartElement("Parameters")
        ' Write out the parameters
        If Me.Parameters IsNot Nothing Then
          For Each lobjParameter As IParameter In Me.Parameters
            ' Write the Parameter element
            .WriteRaw(lobjParameter.ToXmlString)
          Next
        End If
        ' End the Parameters element
        .WriteEndElement()

        ' Open the Operations Element
        .WriteStartElement("Operations")

        ' Write out the operations
        For Each lobjOperable As IOperable In Me.Operations

          ' Write the operation element
          .WriteRaw(lobjOperable.ToXmlElementString)

        Next

        ' End the Operations element
        .WriteEndElement()

        ' Write the RunBeforeBegin
        ' Open the RunBeforeBegin Element
        .WriteStartElement("RunBeforeBegin")

        If RunBeforeBegin IsNot Nothing Then
          ' Write the operable element
          .WriteRaw(RunBeforeBegin.ToXmlElementString)
        End If

        ' End the RunBeforeBegin element
        .WriteEndElement()

        ' Write the RunAfterComplete
        ' Open the RunAfterComplete Element
        .WriteStartElement("RunAfterComplete")

        If RunAfterComplete IsNot Nothing Then
          ' Write the operable element
          .WriteRaw(RunAfterComplete.ToXmlElementString)
        End If

        ' End the RunAfterComplete element
        .WriteEndElement()

        ' Write the RunOnFailure
        ' Open the RunOnFailure Element
        .WriteStartElement("RunOnFailure")

        If RunOnFailure IsNot Nothing Then
          ' Write the operable element
          .WriteRaw(RunOnFailure.ToXmlElementString)
        End If

        ' End the RunOnFailure element
        .WriteEndElement()

        ' Write the RunBeforeParentBegin
        ' Open the RunBeforeParentBegin Element
        .WriteStartElement("RunBeforeParentBegin")

        If RunBeforeParentBegin IsNot Nothing Then
          ' Write the operable element
          .WriteRaw(RunBeforeParentBegin.ToXmlElementString)
        End If

        ' End the RunBeforeParentBegin element
        .WriteEndElement()

        ' Write the RunAfterParentComplete
        ' Open the RunAfterParentComplete Element
        .WriteStartElement("RunAfterParentComplete")

        If RunAfterParentComplete IsNot Nothing Then
          ' Write the operable element
          .WriteRaw(RunAfterParentComplete.ToXmlElementString)
        End If

        ' End the RunAfterParentComplete element
        .WriteEndElement()

        ' Write the RunOnParentFailure
        ' Open the RunOnParentFailure Element
        .WriteStartElement("RunOnParentFailure")

        If RunOnParentFailure IsNot Nothing Then
          ' Write the operable element
          .WriteRaw(RunOnParentFailure.ToXmlElementString)
        End If

        ' End the RunOnParentFailure element
        .WriteEndElement()

        ' For the Job level triggers we will only write them to XML if 
        ' they exist since they are only valid for jobs and not all parents.

        If RunBeforeJobBegin IsNot Nothing Then
          ' Write the RunBeforeJobBegin
          ' Open the RunBeforeJobBegin Element
          .WriteStartElement("RunBeforeJobBegin")

          ' Write the operable element
          .WriteRaw(RunBeforeJobBegin.ToXmlElementString)

          ' End the RunBeforeJobBegin element
          .WriteEndElement()
        End If

        If RunAfterJobComplete IsNot Nothing Then
          ' Write the RunAfterJobComplete
          ' Open the RunAfterJobComplete Element
          .WriteStartElement("RunAfterJobComplete")

          ' Write the operable element
          .WriteRaw(RunAfterJobComplete.ToXmlElementString)

          ' End the RunAfterJobComplete element
          .WriteEndElement()
        End If

        If RunOnJobFailure IsNot Nothing Then
          ' Write the RunOnJobFailure
          ' Open the RunOnJobFailure Element
          .WriteStartElement("RunOnJobFailure")

          ' Write the operable element
          .WriteRaw(RunOnJobFailure.ToXmlElementString)

          ' End the RunOnJobFailure element
          .WriteEndElement()
        End If

      End With

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
      'Return New Process(Me)
      Dim lstrProcessString As String = Me.ToXmlString()
      Return Serializer.Deserialize.XmlString(lstrProcessString, GetType(Process))
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
        ' DISPOSETODO: free other state (managed objects).

        If Me.Operations IsNot Nothing Then
          ' Dispose of each operation
          For Each lobjOperation As IOperable In Me.Operations
            If Not lobjOperation.IsDisposed Then
              lobjOperation.Dispose()
            End If
          Next
        End If

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

  Private Sub MobjRunBeforeJobBegin_Begin(sender As Object, e As OperableEventArgs) Handles MobjRunBeforeJobBegin.Begin
    Try
      ' If e.WorkItem IsNot Nothing AndAlso e.WorkItem.GetType.Name = "JobItemProxy" Then
      If e.WorkItem IsNot Nothing AndAlso TypeOf e.WorkItem Is IJobItemProxy Then
        'Debug.Print(CType(e.WorkItem, Object).Job.Name())
        CType(e.WorkItem, IJobItemProxy).RunBeforeJobBeginCount += 1
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#Region "ILoggable Implementation"

  'Private mobjLogSession As Gurock.SmartInspect.Session

  'Protected Overridable Sub FinalizeLogSession() Implements ILoggable.FinalizeLogSession
  '  Try
  '    ApplicationLogging.FinalizeLogSession(mobjLogSession)
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub

  'Protected Overridable Sub InitializeLogSession() Implements ILoggable.InitializeLogSession
  '  Try
  '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, System.Drawing.Color.MistyRose)
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub

  Private Sub MobjCurrentOperation_OperatingError(sender As Object, e As OperableErrorEventArgs) Handles MobjCurrentOperation.OperatingError
    Try
      If e.Exception IsNot Nothing AndAlso TypeOf e.Exception Is ServiceNotAvailableException Then
        If sender IsNot Nothing AndAlso TypeOf sender Is IOperable Then
          If DirectCast(sender, IOperable).Parent IsNot Nothing Then
            If DirectCast(sender, IOperable).Parent.GetType.Name = "Batch" Then
              sender.Parent.Job.CancelJob(e.Exception.Message)
            End If
          End If
        End If
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub MobjOperations_ItemPropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles MobjOperations.ItemPropertyChanged
    RaiseEvent ParameterPropertyChanged(sender, e)
  End Sub

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

#End Region

End Class