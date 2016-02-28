VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_OptCondParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public OptIdentifier As String
Public AxisPattern As Integer
Public AxisLevel As Double
Public AxisRegisterAxis As Integer
Public Shutter As Integer
Public WedgeFilter As Integer
Public Pattern As Integer
Public color As Integer
Public NDFilter As Integer
Public ColorTemperature As Integer
Public FNumberTurret As Integer
Public OptModFnumberTurret As Integer
Public OptModDownPosition As Integer
Public EPD1 As Integer
Public EPD2 As Integer
Public Slit As Integer
Public LCShutter As Integer
Public Diffusion As Integer
Public FNumberIris As Integer
Public Mirror As Integer
Public LED As Integer
Public Pupil As Integer
Public Illuminant As Integer
Public Frosted As Integer
Public IrisPos As Integer
Public DeviceX As Integer
Public DeviceY As Integer
Public SlideINOUT As Integer
'>>>2012/3/13 M.Imamura Add. for Opt_Judgement
Public OptJudge As String
'<<<2012/3/13 M.Imamura Add. for Opt_Judgement

Private Sub init_cspals_optCondParams()
'クラス初期化
    OptIdentifier = vbNullString
    AxisPattern = -1
    AxisLevel = -1
    AxisRegisterAxis = -1
    Shutter = -1
    WedgeFilter = -1
    Pattern = -1
    color = -1
    NDFilter = -1
    ColorTemperature = -1
    FNumberTurret = -1
    OptModFnumberTurret = -1
    OptModDownPosition = -1
    EPD1 = -1
    EPD2 = -1
    Slit = -1
    LCShutter = -1
    Diffusion = -1
    FNumberIris = -1
    Mirror = -1
    LED = -1
    Pupil = -1
    Illuminant = -1
    Frosted = -1
    IrisPos = -1
    DeviceX = -9999
    DeviceY = -9999
    SlideINOUT = -1
    OptJudge = vbNullString
End Sub

Private Sub Class_Initialize()
    'クラスのインスタンスが生成された際に実行される関数
    '初期化関数の呼び出し
    Call init_cspals_optCondParams
End Sub

