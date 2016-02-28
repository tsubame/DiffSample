VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIdpBitDepth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Enum IdpBitDepth
'    idpDepthS16 = 0
'    idpDepthS32 = 1
'    idpDepthF32 = 2
'End Enum

Private Const mVALID_VALUES As String = "idpDepthS16,idpDepthS32,idpDepthF32"

Dim mSourceCode As String
Dim mValue As IdpBitDepth

Sub SetValue(ByVal pSourceCode As String)
    mSourceCode = Strings.Trim$(pSourceCode)
    Call Compile
End Sub

Private Sub Compile()
    If mSourceCode Like "*S16" Then
        mValue = idpDepthS16
    ElseIf mSourceCode Like "*S32" Then
        mValue = idpDepthS32
    ElseIf mSourceCode Like "*F32" Then
        mValue = idpDepthF32
    Else
        TheError.Raise 9999, TypeName(Me), "Illegal IdpBitDepth: " & mSourceCode
    End If
End Sub

Function GetValue() As IdpBitDepth
    GetValue = mValue
End Function

Function GetSourceCode() As String
    GetSourceCode = mSourceCode
End Function

Function ValidValues() As String
    ValidValues = mVALID_VALUES
End Function
