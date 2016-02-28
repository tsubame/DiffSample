VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDcTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DC測定実行クラスが実装するインターフェイス
'
'目的:
'   各アクション実行時の内容を切り替える
'
'作成者:
'   0145206097
'
Option Explicit

Public Sub SetModeFVMI( _
    ByVal PinList As String, ByVal ForceV As Variant, _
    ByVal ClampI As Double, ByVal SiteIndex As Long, _
    Optional ByVal Connection As Boolean = True _
)
End Sub

Public Sub SetModeFVMIForBPMU( _
    ByVal PinList As String, ByVal ForceV As Variant, _
    ByVal ClampI As Double, ByVal SiteIndex As Long, _
    Optional ByVal Connection As Boolean = True _
)
End Sub
'#V21-Release
Public Sub SetModeFVMIConcurrent( _
    ByVal PinList As String, ByVal ForceV As Variant, _
    ByVal ClampI As Double, _
    Optional ByVal Connection As Boolean = True _
)
End Sub
'#V21-Release
Public Sub SetModeFVMIForBPMUConcurrent( _
    ByVal PinList As String, ByVal ForceV As Variant, _
    ByVal ClampI As Double, _
    Optional ByVal Connection As Boolean = True _
)
End Sub
Public Sub SetModeFIMV( _
    ByVal PinList As String, ByVal ForceI As Variant, _
    ByVal ClampV As Double, ByVal SiteIndex As Long, _
    Optional ByVal Connection As Boolean = True _
)
End Sub

Public Sub SetModeFIMVForBPMU( _
    ByVal PinList As String, ByVal ForceI As Variant, _
    ByVal ClampV As Double, ByVal SiteIndex As Long, _
    Optional ByVal Connection As Boolean = True _
)
End Sub
'#V21-Release
Public Sub SetModeFIMVConcurrent( _
    ByVal PinList As String, ByVal ForceI As Variant, _
    ByVal ClampV As Double, _
    Optional ByVal Connection As Boolean = True _
)
End Sub
'#V21-Release
Public Sub SetModeFIMVForBPMUConcurrent( _
    ByVal PinList As String, ByVal ForceI As Variant, _
    ByVal ClampV As Double, _
    Optional ByVal Connection As Boolean = True _
)
End Sub
Public Sub SetModeMV(ByVal PinList As String, ByVal ClampV As Double, ByVal SiteIndex As Long, _
                     Optional ByVal Connection As Boolean = True)
End Sub

Public Sub SetModeMVForBPMU(ByVal PinList As String, ByVal ClampV As Double, ByVal SiteIndex As Long, _
                          Optional ByVal Connection As Boolean = True)
End Sub
'#V21-Release
Public Sub SetModeMVConcurrent(ByVal PinList As String, ByVal ClampV As Double, _
                     Optional ByVal Connection As Boolean = True)
End Sub
'#V21-Release
Public Sub SetModeMVForBPMUConcurrent(ByVal PinList As String, ByVal ClampV As Double, _
                     Optional ByVal Connection As Boolean = True)
End Sub
Public Sub SetRangeMI(ByVal PinList As String, ByVal ClampI As Double, ByVal SiteIndex As Long)
End Sub

Public Sub SetRangeMIForBPMU(ByVal PinList As String, ByVal ClampI As Double, ByVal SiteIndex As Long)
End Sub
'#V21-Release
Public Sub SetRangeMIConcurrent(ByVal PinList As String, ByVal ClampI As Double)
End Sub
'#V21-Release
Public Sub SetRangeMIForBPMUConcurrent(ByVal PinList As String, ByVal ClampI As Double)
End Sub
Public Sub MeasureVConcurrent(ByVal PinList As String, ByVal avgNum As Long)
End Sub

Public Sub MeasureIConcurrent(ByVal PinList As String, ByVal avgNum As Long)
End Sub

Public Sub MeasureVSerial(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Public Sub MeasureVSerialForBPMU(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Public Sub MeasureVSerialForGANG(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Public Sub MeasureISerial(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Public Sub MeasureISerialForBPMU(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Public Sub MeasureISerialForGANG(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Public Sub Connect(ByVal PinList As String, ByVal SiteIndex As Long)
End Sub

Public Sub ConnectForBPMU(ByVal PinList As String, ByVal SiteIndex As Long)
End Sub

Public Sub Disconnect(ByVal PinList As String, ByVal SiteIndex As Long)
End Sub

Public Sub DisconnectForBPMU(ByVal PinList As String, ByVal SiteIndex As Long)
End Sub
'#V21-Release
Public Sub DisconnectConcurrent(ByVal PinList As String)
End Sub
'#V21-Release
Public Sub DisconnectForBPMUConcurrent(ByVal PinList As String)
End Sub
'#V21-Release
Public Sub LoadResult(ByVal measurePinSet As Collection, Optional ByVal SiteIndex As Long = ALL_SITE)
End Sub

Public Function GetMeasureRange(ByVal PinList As String) As Collection
End Function

Public Function GetMeasureRangeForBPMU(ByVal PinList As String) As Collection
End Function

Public Function GetForceValue(ByVal PinList As String) As Collection
End Function

Public Function GetForceValueForBPMU(ByVal PinList As String) As Collection
End Function

Public Function GetForceVLimit(ByVal PinList As String) As Double()
End Function

Public Function GetForceILimit(ByVal PinList As String) As Double()
End Function

Public Function GetForceVLimitForBPMU(ByVal PinList As String) As Double()
End Function

Public Function GetForceILimitForBPMU(ByVal PinList As String) As Double()
End Function

Public Function GetBoardName(ByVal PinList As String) As String
End Function

Public Function GetBoardNameForBPMU(ByVal PinList As String) As String
End Function

