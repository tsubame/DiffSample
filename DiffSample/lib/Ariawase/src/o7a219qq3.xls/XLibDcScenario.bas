Attribute VB_Name = "XLibDcScenario"
Option Explicit

Public Function CreateCDCScenario() As CDCScenario
    Set CreateCDCScenario = New CDCScenario
End Function

Public Function CreateCStdDCLibV01() As CStdDCLibV01
    Set CreateCStdDCLibV01 = New CStdDCLibV01
End Function

Public Function CreateCPlaybackDc() As CPlaybackDc
    Set CreateCPlaybackDc = New CPlaybackDc
End Function
