Public Interface IDataUpdateStatus

    Property UpdateStatus() As UpdateStatus

End Interface

Public Enum UpdateStatus
    NewRecord = -1
    Updated = 0
    NeedsUpdate = 1
    NeedsRemoving = 2
    Removed = 4
End Enum