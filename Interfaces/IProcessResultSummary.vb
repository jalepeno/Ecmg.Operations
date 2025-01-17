' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IProcessResultSummary.vb
'  Description :  [type_description_here]
'  Created     :  4/26/2016 5:28:40 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Public Interface IProcessResultSummary
  Inherits IOperableResultSummary

  Property VersionCountInfo As IStatistical
  Property FileCountInfo As IStatistical
  Property FileSizeInfo As IStatistical

  Property OperationResults As IOperableResultsSummary

  'Property ProcessingTime As ProcessingTimeInfo

End Interface
