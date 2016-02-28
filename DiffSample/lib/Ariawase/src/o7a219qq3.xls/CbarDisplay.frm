VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CbarDisplay 
   Caption         =   "Cross Reticle Display Start"
   ClientHeight    =   1245
   ClientLeft      =   11040
   ClientTop       =   9330
   ClientWidth     =   3225
   OleObjectBlob   =   "CbarDisplay.frx":0000
End
Attribute VB_Name = "CbarDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private ErrInfo(nSite) As Double

Dim stopEvent As Boolean

Private Sub cmdStart_Click()

       '<Set MSB>
        theidv.MsbRed = 7
        theidv.MsbGreen = 7
        theidv.MsbBlue = 7

       '<IDV Display Size>
        theidv.Zoom (50)
        theidv.MoveImageData 1, 1
        theidv.FormTop = 1
        theidv.FormLeft = 1
        
    Dim acqPlane As CImgPlane
    Dim rawPlane As CImgPlane
    Call GetFreePlane(acqPlane, "acq", idpDepthS16)
    Call GetFreePlane(rawPlane, "vmcu", acqPlane.BitDepth, , "RAW")

    Do
       '<Acquire>
'        Call Capture(2, acqPlane)
'        Call Transfer(ErrInfo)
'        Call PutImageInto(acqPlane, rawPlane, "MIPI_FULL", "FULL_AQ")

       '<Refresh Display>
        theidv.Refresh
        
        TheHdw.WAIT 50 * mS
       
       '<See if stop button click or not!>
        DoEvents

        If stopEvent = True Then
            Unload CbarDisplay
            Exit Sub
        End If

    Loop

End Sub

Private Sub cmdStop_Click()

    stopEvent = True
    Hide
    Unload CbarDisplay

End Sub

