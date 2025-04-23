' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IOperable.vb
'  Description :  [type_description_here]
'  Created     :  11/23/2011 5:08:01 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Globalization
Imports Documents.Core

#End Region

Public Interface IOperable
  Inherits IActionItem
  Inherits IDisposable
  Inherits ICloneable

  ''' <summary>
  ''' The Id of the document to operate on.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Property DocumentId As String

  ' ''' <summary>
  ' ''' The batch with which to execute the item.
  ' ''' </summary>
  ' ''' <value></value>
  ' ''' <returns></returns>
  ' ''' <remarks></remarks>
  'Property Batch As Batch

  Property OperableParent As IOperable

  Property Parent As IItemParent

  ''' <summary>
  ''' The item to execute the operation or process against.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Property WorkItem As IWorkItem

  ''' <summary>
  ''' Used to pass any message regarding the execution of the operation.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Property ProcessedMessage As String

  Property ShouldExecute As Boolean

  Property StartTime() As DateTime

  Property FinishTime() As DateTime

  ReadOnly Property TotalProcessingTime() As TimeSpan

  ''' <summary>
  ''' Used to indicate the result.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ReadOnly Property Result As Result

  ReadOnly Property ResultDetail As IOperableResult

  ''' <summary>
  ''' Gets or sets a value indicating whether or 
  ''' not the result of the operation should be logged.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Property LogResult As Boolean

  ''' <summary>
  ''' Optional property used to provide a reference to the host 
  ''' application or form under which the operation or proces is running.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Property Host As Object

  ''' <summary>
  ''' Optional tag which can be associated with the operation or process.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Property Tag As Object

  ReadOnly Property IsDisposed() As Boolean
  ReadOnly Property Locale As CultureInfo

  Sub CheckParameters()

  Sub UpdateParameterToEnum(lpName As String, lpEnumType As Type)

  'Sub SetDescription(ByVal lpDescription As String)
  Sub SetResult(ByVal lpResult As Result)

  ''' <summary>
  ''' Called to execute.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function Execute(ByRef lpWorkItem As IWorkItem) As Result



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
  'Function ResolveInlineParameter(ByVal lpValue As String) As String

  ''' <summary>
  ''' Raised when an operation starts
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Event Begin As EventHandler(Of OperableEventArgs)

  ''' <summary>
  ''' Raised when an operation completes
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Event Complete As EventHandler(Of OperableEventArgs)

  ''' <summary>
  ''' Raised when an error occurs in an operation.
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Event OperatingError As EventHandler(Of OperableErrorEventArgs)

  Event ParameterPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)

  Function OnExecute() As Result

  ''' <summary>
  ''' Used to reset the operation and initialize the result to not processed.
  ''' </summary>
  ''' <remarks></remarks>
  Sub Reset()
  Sub OnBegin(ByVal e As OperableEventArgs)
  Sub OnComplete(ByVal e As OperableEventArgs)
  Sub OnError(ByVal e As OperableErrorEventArgs)
  Sub SetInstanceId(lpInstanceId As String)

  Property RunBeforeBegin As IOperable
  Property RunAfterComplete As IOperable
  Property RunOnFailure As IOperable
  ReadOnly Property InstanceId As String

  ' <Added by: Ernie at: 12/18/2012-3:35:52 PM on machine: ERNIE-THINK>
  ' Rollback Support
  ReadOnly Property CanRollback As Boolean
  Function Rollback(ByVal lpWorkItem As IWorkItem) As Result
  Function OnRollBack() As Result
  ' </Added by: Ernie at: 12/18/2012-3:35:52 PM on machine: ERNIE-THINK>

  ' ''' <summary>
  ' ''' Gets or sets the collection of operations to execute if this operation fails.
  ' ''' </summary>
  ' ''' <value></value>
  ' ''' <returns></returns>
  ' ''' <remarks></remarks>
  'Property OnFailureOperations As IOperations
  Function ToJson() As String
  Function ToString() As String
  Overloads Function ToXmlElementString() As String

End Interface