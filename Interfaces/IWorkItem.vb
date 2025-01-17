' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IWorkItem.vb
'  Description :  [type_description_here]
'  Created     :  12/2/2011 8:24:27 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

Public Interface IWorkItem

  Property Id() As String
  Property Title() As String
  Property SourceDocId() As String
  Property DestinationDocId() As String
  Property Parent As IItemParent
  Property ProcessedStatus() As ProcessedStatus
  Property Process As IOperable
  Property ProcessResult As IProcessResult
  Property StartTime() As DateTime
  Property FinishTime() As DateTime
  Property TotalProcessingTime() As TimeSpan
  Property ProcessedBy() As String
  Property ProcessedMessage() As String
  Property Document As Document

  Property Folder As Folder

  Property [Object] As CustomObject

  ''' <summary>
  ''' The Tag is a generic property that can be used to store any object.
  ''' It is not used by the system and is available for use by the developer.
  ''' </summary>
  ''' <returns></returns>
  Property Tag As Object

  ReadOnly Property CreateDate As DateTime

  Function Execute(lpProcess As IOperable) As Boolean
  Function ToJsonString(lpIncludeProcessResult As Boolean) As String

End Interface
