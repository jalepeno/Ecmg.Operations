Public Interface IJobItemProxy
  Inherits IWorkItem

  ReadOnly Property JobName() As String
  ReadOnly Property ProjectName() As String
  Property RunBeforeJobBeginCount As Integer

End Interface
