Attribute VB_Name = "PMCInterface"
Public mbOnlineMode As Boolean
Public ng_data(nSite) As Double
Dim g_lngTemperature As String

Option Explicit

Public Function SetOnlineMode(ByRef bOnlineMode As Boolean) As Boolean
    
    mbOnlineMode = bOnlineMode
    SetOnlineMode = True

End Function

Public Function SetFlagShift(ByRef iFlagShift As Integer) As Boolean
    
    Flg_Shift = iFlagShift
    SetFlagShift = True

End Function

Public Function GetNGData(ByRef strNGData() As String, ByRef strUnit() As String) As Boolean

    Dim lSiteNo As Long
    Dim site As Long
    
    For site = 0 To nSite
        strNGData(site) = ng_data(site)

'        strUnit(site) = "mv"

    Next
    
    GetNGData = True
    
End Function

Public Function GetTemperature(ByRef lngTemperature As Long) As Boolean
    
    lngTemperature = TenkenTemp

    GetTemperature = True

End Function

