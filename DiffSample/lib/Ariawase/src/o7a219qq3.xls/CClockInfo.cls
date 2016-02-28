VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CClockInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_DriveHiBase As Double
Private m_DriveLoBase As Double
Private m_ClampHi As Double
Private m_ClampLo As Double
Private m_CompareHi As Double
Private m_CompareLo As Double
Private m_Threshold As Double
Private m_DriveHiOffset() As Double
Private m_DriveLoOffset() As Double


Private Sub Class_Initialize()

    ReDim m_DriveHiOffset(TheExec.sites.ExistingCount - 1)
    ReDim m_DriveLoOffset(TheExec.sites.ExistingCount - 1)
    
End Sub

Public Property Let DriveHiBase(ByVal argVolt As Double)

    m_DriveHiBase = argVolt
    
End Property

Public Property Let DriveHiOffset(ByVal site As Long, ByVal argVolt As Double)

    m_DriveHiOffset(site) = argVolt
        
End Property

Public Property Get DriveHi() As Variant

    Dim tmpVolt() As Double
    Dim site As Long

    ReDim tmpVolt(UBound(m_DriveHiOffset))
    For site = 0 To UBound(tmpVolt)
        tmpVolt(site) = m_DriveHiBase + m_DriveHiOffset(site)
    Next site

    DriveHi = tmpVolt
    
End Property

Public Property Let DriveLoBase(ByVal argVolt As Double)

    m_DriveLoBase = argVolt
    
End Property

Public Property Let DriveLoOffset(ByVal argSite As Long, ByVal argVolt As Double)

    m_DriveLoOffset(argSite) = argVolt
    
End Property

Public Property Get DriveLo() As Variant

    Dim tmpVolt() As Double
    Dim site As Long

    ReDim tmpVolt(UBound(m_DriveLoOffset))
    For site = 0 To UBound(tmpVolt)
        tmpVolt(site) = m_DriveLoBase + m_DriveLoOffset(site)
    Next site

    DriveLo = tmpVolt
    
End Property

Public Property Let ClampHi(ByVal argVolt As Double)

    m_ClampHi = argVolt
    
End Property

Public Property Get ClampHi() As Double

    ClampHi = m_ClampHi
    
End Property

Public Property Let ClampLo(ByVal argVolt As Double)

    m_ClampLo = argVolt
    
End Property

Public Property Get ClampLo() As Double

    ClampLo = m_ClampLo
    
End Property

Public Property Let CompareHi(ByVal argVolt As Double)

    m_CompareHi = argVolt
    
End Property

Public Property Get CompareHi() As Double

    CompareHi = m_CompareHi
    
End Property

Public Property Let CompareLo(ByVal argVolt As Double)

    m_CompareLo = argVolt
    
End Property

Public Property Get CompareLo() As Double

    CompareLo = m_CompareLo
    
End Property

Public Property Let Threshold(ByVal argVolt As Double)

    m_Threshold = argVolt
    
End Property

Public Property Get Threshold() As Double

    Threshold = m_Threshold
    
End Property

Public Sub Force(ByVal argPinName As String)

    Dim site As Long
    Dim locChannels() As Long
    Dim locChanNum As Long
    Dim locSiteNum As Long
    Dim locErrMsg As String
    Dim i As Long
    
    For site = 0 To TheExec.sites.ExistingCount - 1
        If TheExec.sites.site(site).Active = True Then
            Call TheExec.DataManager.GetChanList(argPinName, site, chIO, locChannels, locChanNum, locSiteNum, locErrMsg)
            For i = 0 To UBound(locChannels)
                With TheHdw.PinLevels.chan(locChannels(i))
                    Call .ModifyLevel(chVDriveHi, m_DriveHiBase + m_DriveHiOffset(site))
                    Call .ModifyLevel(chVDriveLo, m_DriveLoBase + m_DriveLoOffset(site))
                    Call .ModifyLevel(chVCH, m_ClampHi)
                    Call .ModifyLevel(chVCL, m_ClampLo)
                    Call .ModifyLevel(chVCompareHi, m_CompareHi)
                    Call .ModifyLevel(chVCompareLo, m_CompareLo)
                    Call .ModifyLevel(chVT, m_Threshold)
                End With
            Next i
        End If
    Next site
    
    Call TheHdw.Digital.relays.Pins(argPinName).ConnectPins
    
End Sub
'#No-Release
Public Sub ForceGroupPins(ByVal argPinName As String)

'警告事項:
'          channel Map上でのサイトシェアがあるピン(グループ)については､
'          この関数は使用しないこと
    
    Dim siteStatus As LoopStatus
    Dim curSite  As Long

    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        'Pass
        curSite = TheExec.sites.SelectedSite
            If TheExec.sites.site(curSite).Active = True Then
                    With TheHdw.PinLevels.Pins(argPinName)
                         Call .ModifyLevel(chVDriveHi, m_DriveHiBase + m_DriveHiOffset(curSite))
                         Call .ModifyLevel(chVDriveLo, m_DriveLoBase + m_DriveLoOffset(curSite))
                         Call .ModifyLevel(chVCH, m_ClampHi)
                         Call .ModifyLevel(chVCL, m_ClampLo)
                         Call .ModifyLevel(chVCompareHi, m_CompareHi)
                         Call .ModifyLevel(chVCompareLo, m_CompareLo)
                         Call .ModifyLevel(chVT, m_Threshold)
                    End With
            End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    'Site Loop End **************************************

    Call TheHdw.Digital.relays.Pins(argPinName).ConnectPins
    
End Sub
'#No-Release
Public Sub ForceGroupPinsMulti(ByVal argPinName As String)

'注意事項:
'          設定値は全サイトがサイト0の設定値固定
'警告事項:
'          channel Map上でのサイトシェアがあるピン(グループ)については､
'          この関数は使用しないこと


                    With TheHdw.PinLevels.Pins(argPinName)
                         Call .ModifyLevel(chVDriveHi, m_DriveHiBase + m_DriveHiOffset(0))
                         Call .ModifyLevel(chVDriveLo, m_DriveLoBase + m_DriveLoOffset(0))
                         Call .ModifyLevel(chVCH, m_ClampHi)
                         Call .ModifyLevel(chVCL, m_ClampLo)
                         Call .ModifyLevel(chVCompareHi, m_CompareHi)
                         Call .ModifyLevel(chVCompareLo, m_CompareLo)
                         Call .ModifyLevel(chVT, m_Threshold)
                    End With

    Call TheHdw.Digital.relays.Pins(argPinName).ConnectPins
    
End Sub
