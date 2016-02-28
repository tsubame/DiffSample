Attribute VB_Name = "Assert"
'''+----                                                                   --+
'''|                             Ariawase 0.6.0                              |
'''|                Ariawase is free library for VBA cowboys.                |
'''|          The Project Page: https://github.com/vbaidiot/Ariawase         |
'''+--                                                                   ----+
Option Explicit
Option Private Module

Private Const TestClassSuffix As String = "Test"
Private Const TestProcSuffix As String = "_Test"

Private Const AssertModule As String = "Assert"
Private Const GeneratedProc As String = "TestRunner"
Private Const CommentLineInGeneratedProc As Long = 1

Private Const ResultLineLen As Long = 76

Private xStartTime As Single
Private xEndTime As Single
Private xSuccSubCount As Long
Private xFailSubCount As Long

Private xAssertIx As Long
Private xFailMsgs As Collection

Private Property Get VBProject() As Object
    Dim app As Object: Set app = Application
    Select Case app.Name
        Case "Microsoft Word":   Set VBProject = app.MacroContainer.VBProject
        Case "Microsoft Excel":  Set VBProject = app.ThisWorkbook.VBProject
        Case "Microsoft Access": Set VBProject = app.VBE.ActiveVBProject
        Case Else: Err.Raise 17
    End Select
End Property

Private Function ProcNames(ByVal vbcompo As Object) As Collection
    Dim cdmdl As Object:     Set cdmdl = vbcompo.CodeModule
    Dim procs As Collection: Set procs = New Collection
    Dim proc As Variant:     proc = ""
    
    Dim i As Long
    For i = 1 + cdmdl.CountOfDeclarationLines To cdmdl.CountOfLines
        If proc <> cdmdl.ProcOfLine(i, 0) Then
            proc = cdmdl.ProcOfLine(i, 0)
            procs.Add proc
        End If
    Next
    
    Set ProcNames = procs
End Function

Private Sub WriteResult(ByVal res As String)
    Debug.Print res
End Sub

Private Sub TestStart(ByVal clsName As String)
    WriteResult String$(ResultLineLen, "-")
    WriteResult clsName
    WriteResult String$(ResultLineLen, "-")
    
    xSuccSubCount = 0
    xFailSubCount = 0
    xStartTime = Timer
End Sub

Private Sub TestEnd()
    xEndTime = Timer
    
    WriteResult String$(ResultLineLen, "=")
    WriteResult _
          xSuccSubCount & " succeeded, " & xFailSubCount & " failed," _
        & " took " & Format$(xEndTime - xStartTime, "0.00") & " seconds."
End Sub

Private Function CheckTestProcName(ByVal proc As String) As Boolean
    CheckTestProcName = Right$(proc, Len(TestProcSuffix)) = TestProcSuffix
End Function

Private Function CheckTestClassName(ByVal clsName As String) As Boolean
    CheckTestClassName = Right$(clsName, Len(TestClassSuffix)) = TestClassSuffix
End Function

Private Sub RunTestSub(ByVal obj As Object, ByVal proc As String)
    xAssertIx = 1
    Set xFailMsgs = New Collection
    
    CallByName obj, proc, VbMethod
    
    If xFailMsgs.Count < 1 Then
        WriteResult "+ " & proc
        IncrPre xSuccSubCount
    Else
        WriteResult "- " & proc
        WriteResult "  " & Join(ClctToArr(xFailMsgs), vbCrLf & "  ")
        IncrPre xFailSubCount
    End If
End Sub

Public Sub RunTestOf(ByVal clsObj As Object)
    Dim clsName As String: clsName = TypeName(clsObj)
    If Not CheckTestClassName(clsName) Then Err.Raise 5
    
    Dim proc As Variant, procs As Collection
    Set procs = ProcNames(VBProject.VBComponents(clsName))
    
    TestStart clsName
    For Each proc In procs
        If CheckTestProcName(proc) Then RunTestSub clsObj, proc
    Next
    TestEnd
End Sub

Public Sub RunTest()
    Call TestRunner
End Sub

Private Sub TestRunner()
    ''' NOTE: This is auto-generated code - don't modify contents of this procedure with the code editor.
End Sub

Public Sub TestRunnerClear()
    Dim asrt As Object: Set asrt = VBProject.VBComponents(AssertModule).CodeModule
    Dim st0 As Long: st0 = asrt.ProcStartLine(GeneratedProc, 0)
    Dim st1 As Long: st1 = asrt.ProcBodyLine(GeneratedProc, 0)
    Dim cnt As Long: cnt = asrt.ProcCountLines(GeneratedProc, 0)
    
    asrt.DeleteLines _
        st1 + (1 + CommentLineInGeneratedProc), _
        cnt - ((st1 - st0) + 2 + CommentLineInGeneratedProc)
End Sub

Public Sub TestRunnerGenerate()
    Dim asrt As Object: Set asrt = VBProject.VBComponents(AssertModule).CodeModule
    Dim st1 As Long: st1 = asrt.ProcBodyLine(GeneratedProc, 0)
    Dim pos As Long: pos = st1 + (1 + CommentLineInGeneratedProc)
    
    Dim vbcompo As Object, ln As String
    For Each vbcompo In VBProject.VBComponents
        If vbcompo.Type = 2 And CheckTestClassName(vbcompo.Name) Then
            ln = "Assert.RunTestOf New " & vbcompo.Name
            asrt.InsertLines pos, vbTab & ln
            IncrPre pos
        End If
    Next
End Sub

Private Sub AssertDone( _
    ByVal isa As Boolean, ByVal cond As Boolean, ByVal msg As String, ByVal exp As Variant, ByVal act As Variant _
    )
    
    If isa <> cond Then
        Push xFailMsgs, "[" & xAssertIx & "] " & msg & ":"
        Push xFailMsgs, "  Expected: " & IIf(isa, "", "Not ") & "<" & Dump(exp) & ">"
        Push xFailMsgs, "  But was:  <" & Dump(act) & ">"
    End If
    IncrPre xAssertIx
End Sub

Public Sub IsNullVal(ByVal x As Variant, Optional ByVal msg As String = "")
    AssertDone True, IsNull(x), msg, Null, x
End Sub

Public Sub IsNotNullVal(ByVal x As Variant, Optional ByVal msg As String = "")
    AssertDone False, IsNull(x), msg, Null, x
End Sub

Public Sub IsInstanceOfTypeName( _
    ByVal expType As String, ByVal x As Variant, Optional ByVal msg As String = "" _
    )
    
    Dim t As String: t = TypeName(x)
    AssertDone True, expType = t, msg, expType, t
End Sub

Public Sub IsNotInstanceOfTypeName( _
    ByVal expType As String, ByVal x As Variant, Optional ByVal msg As String = "" _
    )
    
    Dim t As String: t = TypeName(x)
    AssertDone False, expType = t, msg, expType, t
End Sub

Public Sub AreEq( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone True, Eq(exp, act), msg, exp, act
End Sub

Public Sub AreNotEq( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone False, Eq(exp, act), msg, exp, act
End Sub

Public Sub AreEqual( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone True, Equals(exp, act), msg, exp, act
End Sub

Public Sub AreNotEqual( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone False, Equals(exp, act), msg, exp, act
End Sub

Public Sub AreEqualArr( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone True, ArrEquals(exp, act), msg, exp, act
End Sub

Public Sub AreNotEqualArr( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone False, ArrEquals(exp, act), msg, exp, act
End Sub

Public Sub Fail(Optional ByVal msg As String = "")
    If Len(msg) > 0 Then
        Err.Raise 1004, AssertModule, msg
    Else
        Err.Raise 1004, AssertModule
    End If
End Sub

Public Sub IsErrFunc( _
    ByVal errnum As Variant, _
    ByVal fun As Func, ByVal params As Variant, _
    Optional ByVal msg As String = "" _
    )
    
    If Not (IsEmpty(errnum) Or IsNumeric(errnum)) Then Err.Raise 5
    If Not IsArray(params) Then Err.Raise 5
    
    On Error GoTo Catch
    
    Dim act As Variant: act = Empty
    
    Dim buf As Variant, ret As Boolean
    fun.CallByPtr buf, params
    AssertDone True, ret, msg, errnum, act
    GoTo Escape
    
Catch:
    act = Err.Number
    ret = IsEmpty(errnum) Or act = errnum
    Resume Next
    
Escape:
End Sub

Public Sub IsErrMethod( _
    ByVal errnum As Variant, _
    ByVal obj As Object, ByVal proc As String, ByVal params As Variant, _
    Optional ByVal msg As String = "" _
    )
    
    If Not (IsEmpty(errnum) Or IsNumeric(errnum)) Then Err.Raise 5
    If Not IsArray(params) Then Err.Raise 5
    If LBound(params) <> 0 Then Err.Raise 5
    
    On Error GoTo Catch
    
    Dim act As Variant: act = Empty
    
    Dim i As Long, ret As Boolean
    Dim ubParam As Long: ubParam = UBound(params)
    Dim ps() As Variant: ReDim ps(ubParam)
    For i = 0 To ubParam
        If IsObject(params(i)) Then
            Set ps(i) = params(i)
        Else
            Let ps(i) = params(i)
        End If
    Next
    rtcCallByName obj, StrPtr(proc), VbMethod, ps
    
    AssertDone True, ret, msg, errnum, act
    GoTo Escape
    
Catch:
    act = Err.Number
    ret = IsEmpty(errnum) Or act = errnum
    Resume Next
    
Escape:
End Sub
