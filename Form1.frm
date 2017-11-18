VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '單線固定
   Caption         =   "WinCap"
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8895
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  '沒有框線
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
   Begin VB.Menu mConnect 
      Caption         =   "Connect(&A)"
   End
   Begin VB.Menu mDisconnect 
      Caption         =   "Disconnect(&B)"
   End
   Begin VB.Menu mOption 
      Caption         =   "Option(&C)"
   End
   Begin VB.Menu mSave 
      Caption         =   "SavePicture(&D)"
   End
   Begin VB.Menu mClip 
      Caption         =   "ToClipboard(&E)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hwdc As Long
Dim startcap As Boolean
Dim temp As Long

Private Sub Form_Unload(Cancel As Integer)
    temp = SendMessage(hwdc, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
End Sub

Private Sub mClip_Click()
SendMessage hwdc, 1054, 0, 0 ' WM_CAP_EDIT_COPY
End Sub

Private Sub mConnect_Click()
    hwdc = capCreateCaptureWindow("Dixanta Vision System", ws_child Or ws_visible, 0, 0, 640, 480, Picture1.hWnd, 0)
    If (hwdc <> 0) Then
        temp = SendMessage(hwdc, wm_cap_driver_connect, 0, 0)
        temp = SendMessage(hwdc, wm_cap_set_preview, 1, 0)
        temp = SendMessage(hwdc, WM_CAP_SET_PREVIEWRATE, 30, 0)
        startcap = True
        Else
        MsgBox ("No Webcam found")
    End If
End Sub

Private Sub mDisconnect_Click()
    If startcap = True Then
        temp = SendMessage(hwdc, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
        startcap = False
    End If
End Sub

Private Sub mOption_Click()

    If startcap = True Then
        temp = SendMessage(hwdc, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
    End If
End Sub

Private Sub mSave_Click()
    SendMessage hwdc, 1054, 0, 0 ' WM_CAP_EDIT_COPY
    SavePicture Clipboard.GetData(2), App.Path & "\" & Format(Now, "yyyyMMddHHmmss") & ".bmp"
End Sub
