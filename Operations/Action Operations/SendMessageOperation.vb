' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  SendMessageOperation.vb
'  Description :  [type_description_here]
'  Created     :  8/17/2012 1:57:38 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities
Imports Documents.Messaging

#End Region

Public Class SendMessageOperation
  Inherits ActionOperation

#Region "Class Constants"

  Private Const OPERATION_NAME As String = "SendMessage"
  Friend Const PARAM_MESSAGING_SERVICE As String = "Service"
  Friend Const PARAM_MESSAGING_SERVICE_VERSION As String = "ServiceVersion"
  Friend Const PARAM_FROM_ADDRESS As String = "FromAddress"
  Friend Const PARAM_SUBJECT As String = "Subject"

  Friend Const PARAM_TO_RECIPIENTS As String = "ToRecipients"
  Friend Const PARAM_CC_RECIPIENTS As String = "CcRecipients"
  Friend Const PARAM_BCC_RECIPIENTS As String = "BccRecipients"

  Private Const SERVICE_TYPE_EXCHANGE As String = "Exchange"
  Private Const SERVICE_TYPE_SMTP As String = "Smtp"

  Private Const EXCHANGE_VERSION_2007_SP1 As String = "Exchange2007_SP1"
  Private Const EXCHANGE_VERSION_2010 As String = "Exchange2010"
  Private Const EXCHANGE_VERSION_2010_SP1 As String = "Exchange2010_SP1"
  Private Const EXCHANGE_VERSION_2010_SP2 As String = "Exchange2010_SP2"

#End Region

#Region "Class Variables"

  Private mobjMessagingService As IMessagingService = Nothing
  Private mobjMessage As IMessage = Nothing
  Private mstrSubject As String = String.Empty
  Private mstrFromAddress As String = String.Empty
  Private mobjToRecipients As New List(Of String)
  Private mobjCcRecipients As New List(Of String)
  Private mobjBccRecipients As New List(Of String)
  Private mstrBody As String = String.Empty

#End Region

#Region "Public Properties"

  Public Property Subject As String
    Get
      Try
        Return mstrSubject
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As String)
      Try
        mstrSubject = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Body As String
    Get
      Try
        Return mstrBody
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As String)
      Try
        mstrBody = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property FromAddress As String
    Get
      Try
        Return mstrFromAddress
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(value As String)
      Try
        mstrFromAddress = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public ReadOnly Property ToRecipients As List(Of String)
    Get
      Try
        Return mobjToRecipients
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property CcRecipients As List(Of String)
    Get
      Try
        Return mobjCcRecipients
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property BccRecipients As List(Of String)
    Get
      Try
        Return mobjBccRecipients
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

#End Region

#Region "Public Overrides Methods"

  Public Overrides ReadOnly Property Name As String
    Get
      Try
        Return OPERATION_NAME
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Overrides ReadOnly Property CanRollback As Boolean
    Get
      Return False
    End Get
  End Property

  Friend Overrides Function OnExecute() As OperationEnumerations.Result
    Try
      Return SendMessage()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Protected Methods"

  Protected Overrides Function GetDefaultParameters() As IParameters
    Try

      Dim lobjParameters As IParameters = New Parameters

      If lobjParameters.Contains(PARAM_MESSAGING_SERVICE) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_MESSAGING_SERVICE,
          SERVICE_TYPE_EXCHANGE, "The messaging service to use for sending the notifications."))
      End If

      If lobjParameters.Contains(PARAM_MESSAGING_SERVICE_VERSION) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_MESSAGING_SERVICE_VERSION,
          EXCHANGE_VERSION_2007_SP1, "When using the Exchange Web Service, this parameter specifies which service version to use."))
      End If

      If lobjParameters.Contains(PARAM_FROM_ADDRESS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_FROM_ADDRESS, String.Empty,
          "The from address for email notifications."))
      End If

      If lobjParameters.Contains(PARAM_SUBJECT) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_SUBJECT, String.Empty,
          "The subject for email notifications."))
      End If

      If lobjParameters.Contains(PARAM_TO_RECIPIENTS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_TO_RECIPIENTS, String.Empty,
          "The to addresses for email notifications."))
      End If

      If lobjParameters.Contains(PARAM_CC_RECIPIENTS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_CC_RECIPIENTS, String.Empty,
          "The CC addresses for email notifications."))
      End If

      If lobjParameters.Contains(PARAM_BCC_RECIPIENTS) = False Then
        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_BCC_RECIPIENTS, String.Empty,
          "The BCC addresses for email notifications."))
      End If

      Return lobjParameters

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Function SendMessage() As Result
    Try
      Dim lobjMessage As IMessage
      Dim lstrMessagingService As String = GetStringParameterValue(PARAM_MESSAGING_SERVICE, SERVICE_TYPE_EXCHANGE)
      Dim lstrMessagingServiceVersion As String = GetStringParameterValue(PARAM_MESSAGING_SERVICE_VERSION, EXCHANGE_VERSION_2007_SP1)
      Dim lobjMessagingService As IMessagingService

      If String.IsNullOrEmpty(Me.Subject) Then
        Me.Subject = GetStringParameterValue(PARAM_SUBJECT, "Process Message")
      End If

      If String.IsNullOrEmpty(Me.Body) Then
        '  Me.Body = GetParameterValue(PARAM_SUBJEC, "Process Message")
        'If Me.WorkItem.GetType.Name = "JobItemProxy" Then
        If TypeOf Me.WorkItem Is IJobItemProxy Then
          Dim lstrParentName As String = CType(Me.WorkItem, IJobItemProxy).JobName
          Dim lstrProjectName As String = CType(Me.WorkItem, IJobItemProxy).ProjectName
          If Helper.CallStackContainsMethodName("BeforeJobBegin") Then
            Me.Body = String.Format("Starting Job '{0}' in Project '{1}'", lstrParentName, lstrProjectName)
            Me.Subject = Me.Body
          ElseIf Helper.CallStackContainsMethodName("AfterJobComplete") Then
            Me.Body = String.Format("Completed Job '{0}' in Project '{1}'", lstrParentName, lstrProjectName)
            Me.Subject = Me.Body
          ElseIf Helper.CallStackContainsMethodName("JobError") Then
            Me.Body = String.Format("Error in Job {0}", lstrParentName)
            Me.Subject = Me.Body
          End If
          ' ElseIf Me.WorkItem.GetType.Name = "BatchItemProxy" Then
        ElseIf TypeOf Me.WorkItem Is IBatchItemProxy Then
          Dim lstrParentName As String = CType(Me.WorkItem, IBatchItemProxy).JobName
          If Helper.CallStackContainsMethodName("RunBeforeParentBegin") Then
            Me.Body = String.Format("Starting Job {0}", lstrParentName)
          ElseIf Helper.CallStackContainsMethodName("RunAfterParentComplete") Then
            Me.Body = String.Format("Completed Job {0}", lstrParentName)
            'ElseIf Helper.CallStackContainsMethodName("JobError") Then
            '  Me.Body = String.Format("Error in Job {0}", lstrParentName)
          End If

        End If
      End If

      If String.IsNullOrEmpty(Me.FromAddress) Then
        Me.FromAddress = GetStringParameterValue(PARAM_FROM_ADDRESS, String.Empty)
      End If


      GetAllRecipientsFromParameters()

      Select Case lstrMessagingService
        Case SERVICE_TYPE_EXCHANGE
          Select Case lstrMessagingServiceVersion
            Case EXCHANGE_VERSION_2007_SP1
              lobjMessagingService = New ExchangeMessagingService(ExchangeMessagingService.ExchangeVersion.Exchange2007_SP1, FromAddress)
            Case EXCHANGE_VERSION_2010
              lobjMessagingService = New ExchangeMessagingService(ExchangeMessagingService.ExchangeVersion.Exchange2010, FromAddress)
            Case EXCHANGE_VERSION_2010_SP1
              lobjMessagingService = New ExchangeMessagingService(ExchangeMessagingService.ExchangeVersion.Exchange2010_SP1, FromAddress)
            Case EXCHANGE_VERSION_2010_SP2
              lobjMessagingService = New ExchangeMessagingService(ExchangeMessagingService.ExchangeVersion.Exchange2010_SP2, FromAddress)
            Case Else
              lobjMessagingService = New ExchangeMessagingService(ExchangeMessagingService.ExchangeVersion.Exchange2007_SP1, FromAddress)
          End Select

          If lobjMessagingService IsNot Nothing Then
            lobjMessage = New Message(lobjMessagingService)
            With lobjMessage
              .Subject = Me.Subject
              .Body = Me.Body
              If Not String.IsNullOrEmpty(Me.FromAddress) Then
                .FromAddress = Me.FromAddress
              End If
              .ToRecipients = Me.ToRecipients
              .CcRecipients = Me.CcRecipients
              .BccRecipients = Me.BccRecipients
              .Send()
            End With
          End If

          menuResult = OperationEnumerations.Result.Success

        Case Else
          Me.ProcessedMessage = String.Format("Messaging service '{0}' not yet supported.", lstrMessagingService)
          menuResult = OperationEnumerations.Result.Failed
          OnError(New OperableErrorEventArgs(Me, WorkItem, Me.ProcessedMessage))
      End Select

      Return menuResult

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      menuResult = OperationEnumerations.Result.Failed
      OnError(New OperableErrorEventArgs(Me, WorkItem, ex))
      Return menuResult
    End Try
  End Function

  Protected Sub GetAllRecipientsFromParameters()
    Try
      InitializeRecipientsFromParameter(Parameters.Item(PARAM_TO_RECIPIENTS), ToRecipients)
      InitializeRecipientsFromParameter(Parameters.Item(PARAM_CC_RECIPIENTS), CcRecipients)
      InitializeRecipientsFromParameter(Parameters.Item(PARAM_BCC_RECIPIENTS), BccRecipients)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Protected Sub InitializeRecipientsFromParameter(parameter As IParameter, ByRef recipients As List(Of String))
    Try
      Select Case parameter.Name

        Case PARAM_TO_RECIPIENTS, PARAM_CC_RECIPIENTS, PARAM_BCC_RECIPIENTS
          If parameter.HasValue = False Then
            'Throw New Exceptions.ParameterValueNotSetException(parameter.Name)
            Exit Sub
          End If
          Dim lstrRecipients As String() = parameter.Value.ToString.Split(CChar(";"))
          recipients.Clear()
          For Each lstrRecipient As String In lstrRecipients
            If lstrRecipient.Contains("@") Then
              recipients.Add(lstrRecipient)
            End If
          Next

        Case Else
          Throw New ArgumentOutOfRangeException("parameter",
            String.Format("Parameter '{0}' was not one of the expected parameters.", parameter.Name))
      End Select

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class
