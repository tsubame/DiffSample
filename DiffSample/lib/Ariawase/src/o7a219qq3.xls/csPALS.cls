VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' クラス名  ：  csPALS.cls
' 概要      ：  PALS全体を制御するクラス
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/09/01　新規作成        K.Sumiyashiki
'==========================================================================

Public CommonInfo As csPALS_Common      'csPALS_Commonクラスの定義

Public LoopParams As csPALS_LoopMain    'csPALS_LoopMainクラスの定義


'*********************************************************************************
' 関数名    ：  Class_Initialize
' 概要      ：  インスタンス生成時の初期化
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/09/01　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()

On Error GoTo errPALSClass_Initialize

    Set CommonInfo = New csPALS_Common
        
    Set LoopParams = New csPALS_LoopMain

Exit Sub

errPALSClass_Initialize:
    Call sub_errPALS("csPALS not formed at 'Class_Initialize'", "0-4-01-0-13")

End Sub


'*********************************************************************************
' 関数名    ：  Class_Terminate
' 概要      ：  インスタンス開放時の処理
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/09/01　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Terminate()
    Set CommonInfo = Nothing
    Set LoopParams = Nothing
End Sub

