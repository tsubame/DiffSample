Attribute VB_Name = "XLibReadWrite"
Option Explicit

Public Function CreateCDataSheetManager() As CDataSheetManager
    Set CreateCDataSheetManager = New CDataSheetManager
End Function

Public Function CreateCDcScenarioSheetReader() As CDcScenarioSheetReader
    Set CreateCDcScenarioSheetReader = New CDcScenarioSheetReader
End Function

Public Function CreateCDcScenarioSheetWriter() As CDcScenarioSheetWriter
    Set CreateCDcScenarioSheetWriter = New CDcScenarioSheetWriter
End Function

Public Function CreateCInstanceSheetReader() As CInstanceSheetReader
    Set CreateCInstanceSheetReader = New CInstanceSheetReader
End Function

Public Function CreateCJobListSheetReader() As CJobListSheetReader
    Set CreateCJobListSheetReader = New CJobListSheetReader
End Function

Public Function CreateCDcScenarioSheetLogWriter() As CDcScenarioSheetLogWriter
    Set CreateCDcScenarioSheetLogWriter = New CDcScenarioSheetLogWriter
End Function

Public Function CreateCDcTextFileLogWriter() As CDcTextFileLogWriter
    Set CreateCDcTextFileLogWriter = New CDcTextFileLogWriter
End Function

Public Function CreateCDcPlaybackSheetReader() As CDcPlaybackSheetReader
    Set CreateCDcPlaybackSheetReader = New CDcPlaybackSheetReader
End Function

Public Function CreateCDcPlaybackSheetWriter() As CDcPlaybackSheetWriter
    Set CreateCDcPlaybackSheetWriter = New CDcPlaybackSheetWriter
End Function

Public Function CreateCOffsetSheetReader() As COffsetSheetReader
    Set CreateCOffsetSheetReader = New COffsetSheetReader
End Function
