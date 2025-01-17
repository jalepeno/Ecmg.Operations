' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  OperationEnumerations.vb
'  Description :  [type_description_here]
'  Created     :  11/18/2011 3:15:23 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Public Module OperationEnumerations

  ''' <summary>
  ''' Used to indicate the scope of the operation.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum OperationScope

    ''' <summary>
    ''' Indicates that the operation is on the source document.
    ''' </summary>
    ''' <remarks></remarks>
    Source

    ''' <summary>
    ''' Indicates that the operation is on the destination document.
    ''' </summary>
    ''' <remarks></remarks>
    Destination

  End Enum

  ''' <summary>
  ''' Used to indicate the status of an operation or process.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum ProcessedStatus
    NotProcessed = 0
    Success = 1
    Failed = 2
    Processing = 3
    PreviouslySucceeded = -1
    PreviouslyFailed = -2
  End Enum

  ''' <summary>
  ''' Used to indicate the result of the operation(s).
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum Result

    ''' <summary>
    ''' Indicates that the operation succeeded.
    ''' </summary>
    ''' <remarks></remarks>
    Success = -1

    ''' <summary>
    ''' Indicates that the operation failed.
    ''' </summary>
    ''' <remarks></remarks>
    Failed = 0

    ''' <summary>
    ''' Indicates the operation has not yet started.
    ''' </summary>
    ''' <remarks></remarks>
    NotProcessed = 1

    ''' <summary>
    ''' Indicates that the operation already succeeded and did not get reprocessed.
    ''' </summary>
    ''' <remarks></remarks>
    PreviouslySucceeded = -101

    ''' <summary>
    ''' Indicates that the operation already failed and did not get reprocessed.
    ''' </summary>
    ''' <remarks></remarks>
    PreviouslyFailed = -100

    ''' <summary>
    ''' Indicates that the operation could not be rolled back 
    ''' because rollback is not supported for the operation.    
    ''' </summary>
    ''' <remarks></remarks>
    RollbackNotSupported = -2

    ''' <summary>
    ''' Indicates that the rollback succeeded.
    ''' </summary>
    ''' <remarks></remarks>
    RollbackSuccess = -11

    ''' <summary>
    ''' Indicates that the rollback failed.
    ''' </summary>
    ''' <remarks></remarks>
    RollbackFailed = -10

    ''' <summary>
    ''' Indicates that the operation could not be rolled back 
    ''' because no rollback implementation is available.    
    ''' </summary>
    ''' <remarks></remarks>
    RollbackNotImplemented = -12

  End Enum

End Module