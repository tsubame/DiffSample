VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   画像の状況を返すクラス
'目的:
'   Interceptorで画像状況をログに落とす
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月6日

Option Explicit

Private m_Count As Long             'NowState 出力カウンタ
Private m_LogCollection As Collection                               'ログ格納用
Private m_SaveLogFile As CFileOut                                    'ログファイル入出力制御用インスタンス

Private m_AcqtblArg0 As String
Private m_AcqtblArg1 As String

Private Const HEADER_IMMEDIATE As String = _
                      "PlaneName," _
                    & "Group," _
                    & "UseStatus," _
                    & "GetStatus," _
                    & "Get Retry Times," _
                    & "CurrentPMD," _
                    & "Comment," _
                    & "BankName," _
                    & "Flag," _
                    & "bit1,bit2,bit3,bit4,bit5,bit6,bit7,bit8,bit9,bit10,bit11,bit12,bit13,bit14,bit15,bit16"

'CImgPlaneList.NowState ログ出力のヘッダ。ImgPlaneManager、ImgFlag、ImgPlaneBankで使用
Private Const HEADER_LOG_STRING As String = _
                                      "[RecordNumber]," _
                                    & "[RecordDate]," _
                                    & "[RecordTime]," _
                                    & "[TestInstanceName]," _
                                    & "[ExecuteMethod]," _
                                    & "[ACQTBL_Arg0]," _
                                    & "[ACQTBL_Arg1]," _
                                    & "[GroupName]," _
                                    & "[PlaneName]," _
                                    & "[PlaneMapName]," _
                                    & "[BitDepth]," _
                                    & "[GetStatus]," _
                                    & "[GetRetryTimes]," _
                                    & "[BankName]," _
                                    & "[BankOverWriteMode]," _
                                    & "[Flag]," _
                                    & "[bit]," _
                                    & "[Comment]"

Public Property Let ACQTBL_Arg0(pACQTBL_Arg0 As String)
    m_AcqtblArg0 = pACQTBL_Arg0
End Property

Public Property Get ACQTBL_Arg0() As String
    ACQTBL_Arg0 = m_AcqtblArg0
End Property

Public Property Let ACQTBL_Arg1(pACQTBL_Arg1 As String)
    m_AcqtblArg1 = pACQTBL_Arg1
End Property

Public Property Get ACQTBL_Arg1() As String
    ACQTBL_Arg1 = m_AcqtblArg1
End Property

Public Function GetUsing() As String
    GetUsing = RetImgInfo("Using")
End Function

Public Function GetPlaneBank() As String
    GetPlaneBank = RetImgInfo("PlaneBank")
End Function

Public Function GetFlag() As String
    GetFlag = RetImgInfo("Flag")
End Function

Public Function GetAll() As String
    GetAll = RetImgInfo("All")
End Function

Public Sub Clear()
    
    Dim RetFileName   As String
    Dim RetJobName   As String
    Dim i As Long
    
    If TheIDP.SaveMode = True Or m_LogCollection.Count > 0 Then
        RetFileName = TheExec.ExcelHandle.ActiveWorkbook.Path & "\" & TheExec.ExcelHandle.ActiveWorkbook.Name
        RetJobName = TheExec.CurrentJob                             'JobList の JobName より。無ければファイル名
        
        Set m_SaveLogFile = New CFileOut
        Call m_SaveLogFile.SetFileName(TheIDP.saveFileName)
        
        Call m_SaveLogFile.OpenFile
        Print #m_SaveLogFile.fileNum, "[Test Program],[" & RetFileName & "]"
        Print #m_SaveLogFile.fileNum, "[Job],[" & RetJobName & "]"
    
            Print #m_SaveLogFile.fileNum, HEADER_LOG_STRING
            For i = 1 To m_LogCollection.Count
                Print #m_SaveLogFile.fileNum, m_LogCollection.Item(CStr(i))
            Next i
        
        Call m_SaveLogFile.CloseFile
    End If
    
    Set m_LogCollection = New Collection
    m_Count = 0

End Sub

Public Function NowState(MethodName As String, _
                        Optional PlaneName As String = "-", _
                        Optional PlaneMapName As String = "-", _
                        Optional group As String = "-", _
                        Optional BitDepth As IdpBitDepth = 99, _
                        Optional BankName As String = "-", _
                        Optional BankRetryFlag As String = "-", _
                        Optional GetStatus As String = "-", _
                        Optional GetRetryTimes As String = "-", _
                        Optional flag As String = "-", _
                        Optional Bit As String = "-", _
                        Optional Comment As String = "-" _
                        ) As String

    Dim p_InstanceName As String

    m_Count = m_Count + 1
    
    If TheExec.Flow.IsRunning Then
        p_InstanceName = TheExec.DataManager.InstanceName
    Else
        p_InstanceName = "-"
    End If
    
    Call m_LogCollection.Add(m_Count _
                & "," & CStr(Date) _
                & "," & CStr(TIME) _
                & "," & p_InstanceName _
                & "," & MethodName _
                & "," & TheIDP.PlaneList.ACQTBL_Arg0 _
                & "," & TheIDP.PlaneList.ACQTBL_Arg1 _
                & "," & group _
                & "," & PlaneName _
                & "," & PlaneMapName _
                & "," & idpDepthToStr(BitDepth) _
                & "," & GetStatus _
                & "," & GetRetryTimes _
                & "," & BankName _
                & "," & BankRetryFlag _
                & "," & flag _
                & "," & Bit _
                & "," & Comment _
                , CStr(m_Count))

End Function

Private Function RetImgInfo(ByVal status As String) As String

    Dim pString As String
    Dim pImgInfo As String

    '*****Using Plane List *******************************************************************
    Dim arrUsingPlanes As Collection
    Set arrUsingPlanes = TheIDP.DumpUsingPlane
    
    RetImgInfo = HEADER_IMMEDIATE & vbCrLf
    
    If arrUsingPlanes.Count > 0 Then
        'PlaneBankからリストを取得
        Dim PlaneList As Variant
        If TheIDP.PlaneBank.Count > 0 Then
            PlaneList = Split(Replace(TheIDP.PlaneBank.List, vbCrLf, ","), ",")
        Else
            Set PlaneList = Nothing
        End If
        
        Dim p As CImgPlane
        For Each p In arrUsingPlanes
            pImgInfo = AttributeImg(p, PlaneList)
            If (pImgInfo Like (status & "\\*")) Or (status = "All") Or (status = "Using") Then
                If pImgInfo = "" Then
                    pString = pString & p.Name & "," _
                                & p.planeGroup & "," _
                                & "Use," _
                                & "-," _
                                & "-," _
                                & p.CurrentPmdName & "," _
                                & p.Comment & "," _
                                & "-," _
                                & "-," _
                                & "-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-" _
                                & vbCrLf
                ElseIf pImgInfo Like "PlaneBank\\*" Then
                    pString = pString & p.Name & "," _
                                & p.planeGroup & "," _
                                & "Use," _
                                & "-," _
                                & "-," _
                                & p.CurrentPmdName & "," _
                                & p.Comment & "," _
                                & Replace(pImgInfo, "PlaneBank\\", "") & "," _
                                & "-," _
                                & "-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-" _
                                & vbCrLf
                ElseIf pImgInfo Like "Flag\\*" Then
                    pString = pString & p.Name & "," _
                                & p.planeGroup & "," _
                                & "Use," _
                                & "-," _
                                & "-," _
                                & p.CurrentPmdName & "," _
                                & p.Comment & "," _
                                & "-," _
                                & Replace(pImgInfo, "\\", "") _
                                & vbCrLf
                Else
                
                End If
            End If
        Next p
        '##### 使っているプレーンの情報を返す #####
        RetImgInfo = RetImgInfo & pString
    End If


    '***** No Use Image List *****************************************************************
    Dim Base As Variant
    Dim bases As Collection
    Set bases = TheIDP.DumpPlaneGroup
    'プレングループがなければ終了
    If bases.Count = 0 Then Exit Function
    
    pString = ""
    If status = "All" Then
        Dim NoUsePlane As Variant
        Dim NoUsePlanes As Collection
        Set NoUsePlanes = New Collection
        
        For Each Base In bases
            Call TheIDP.PlaneManager(Base).GetNoUsePlanes(NoUsePlanes)
            For Each NoUsePlane In NoUsePlanes
                pString = pString & CStr(NoUsePlane) & "," & CStr(Base) _
                        & ",No Use" _
                        & ",-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-" _
                        & vbCrLf
            Next NoUsePlane
            Set NoUsePlanes = Nothing
            Set NoUsePlanes = New Collection
        Next Base
    End If
    
    '##### 使っていないプレーンの情報を返す #####
    RetImgInfo = RetImgInfo & pString

End Function
    
'#FlagExpansion
Private Function AttributeImg(ByVal plane As CImgPlane, ByRef PlaneList As Variant) As String
'内容:
'   画像の属性を返す
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月6日
'パラメータ:
'   [plane] In          1):
'   [PlaneList] In      2):
'戻り値:
'   String
'
'注意事項:
'

    '##### PlaneBankかどうか #####
    Dim i As Long
    If IsArray(PlaneList) Then                      'PlaneList にBank の情報が入っていれば配列
        For i = 1 To UBound(PlaneList) Step 2
            If plane.Name = PlaneList(i) Then
                AttributeImg = "PlaneBank\\" & PlaneList(i - 1)
                Exit Function
            End If
        Next i
    End If
    
    '##### SharedFlagかどうか #####
    Dim FlagPlanes As Collection
    Set FlagPlanes = plane.Manager.GetSharedFlagPlanes
    Dim FlagPlane As IImgFlag
    
    For Each FlagPlane In FlagPlanes
        If plane.Name = FlagPlane.Name Then
            AttributeImg = "Flag\\" & GetFlagList(FlagPlane)
            Exit Function
        End If
    Next
    
    AttributeImg = ""
    
End Function

'#FlagExpansion
Private Function GetFlagList(ByRef pFlagPlane As IImgFlag) As String
'bit1,bit2,bit3,bit4,bit5,,,,,,,
'

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim flagName As String

    If pFlagPlane.FlgPlane.BitDepth = idpDepthS16 Then k = 16
    
    'FlagListを取得し配列化
    Dim FlagList As Variant
    FlagList = Split(Replace(pFlagPlane.FlagBitList, vbCrLf, ":"), ":")

    For i = 1 To k
        For j = 0 To UBound(FlagList) - 1 Step 2
            If CStr(i) = FlagList(j) Then flagName = FlagList(j + 1)
        Next j
        If flagName = "" Then
            GetFlagList = GetFlagList & ",-"
        Else
            GetFlagList = GetFlagList & "," & flagName
        End If
        flagName = ""
    Next i

End Function

Private Function idpDepthToStr(ByVal idpDepth As IdpBitDepth) As String
    Select Case idpDepth
        Case idpDepthF32
            idpDepthToStr = "idpDepthF32"
        Case idpDepthS16
            idpDepthToStr = "idpDepthS16"
        Case idpDepthS32
            idpDepthToStr = "idpDepthS32"
        Case Else
            idpDepthToStr = "-"
    End Select
End Function

Private Sub Class_Initialize()
    m_Count = 0
    Set m_LogCollection = New Collection
    Set m_SaveLogFile = Nothing
    
    m_AcqtblArg0 = "-"
    m_AcqtblArg1 = "-"
End Sub

Private Sub Class_Terminate()
    m_Count = 0
    Set m_LogCollection = Nothing
    Set m_SaveLogFile = Nothing
End Sub


