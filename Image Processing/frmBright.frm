VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBright 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Brightness"
   ClientHeight    =   1260
   ClientLeft      =   1905
   ClientTop       =   1860
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   84
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1020
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3180
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin MSComctlLib.Slider sldBright 
      Height          =   255
      Left            =   1020
      TabIndex        =   3
      Top             =   120
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   51
      Min             =   -255
      Max             =   255
      TickFrequency   =   51
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Previewing"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   780
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brightness"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-255"
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   315
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      Height          =   195
      Index           =   4
      Left            =   5400
      TabIndex        =   4
      Top             =   480
      Width           =   270
   End
   Begin VB.Line lin 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   376
      X2              =   8
      Y1              =   47
      Y2              =   47
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   0
      X1              =   376
      X2              =   8
      Y1              =   48
      Y2              =   48
   End
End
Attribute VB_Name = "frmBright"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vBright As Long
Public Preview As Boolean

Private Sub cmdCancel_Click()
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCSour(mdiImgProcess.ActiveForm.Tag), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Preview = False
    Screen.MousePointer = vbHourglass
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    Preview = True
    Screen.MousePointer = vbHourglass
    Call Brightness(mdiImgProcess.ActiveForm.Tag)
    Screen.MousePointer = vbDefault
    ProgressBar.Value = 0
End Sub

Private Sub sldBright_Change()
    vBright = sldBright.Value
End Sub

Private Sub sldBright_Scroll()
    vBright = sldBright.Value
End Sub
