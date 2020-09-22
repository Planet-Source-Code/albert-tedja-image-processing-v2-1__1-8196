VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RGB Color Balance"
   ClientHeight    =   2160
   ClientLeft      =   2160
   ClientTop       =   2115
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   4380
      TabIndex        =   10
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1860
      TabIndex        =   8
      Top             =   1740
      Width           =   1215
   End
   Begin MSComctlLib.Slider sldRed 
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   180
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   25
      Min             =   -100
      Max             =   100
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider sldGreen 
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   25
      Min             =   -100
      Max             =   100
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider sldBlue 
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1020
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   25
      Min             =   -100
      Max             =   100
      TickFrequency   =   25
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Previewing"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   780
   End
   Begin VB.Line lin 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   372
      X2              =   12
      Y1              =   107
      Y2              =   107
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   0
      X1              =   372
      X2              =   12
      Y1              =   108
      Y2              =   108
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   195
      Index           =   4
      Left            =   5340
      TabIndex        =   7
      Top             =   1380
      Width           =   270
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-100"
      Height          =   195
      Index           =   3
      Left            =   900
      TabIndex        =   6
      Top             =   1380
      Width           =   315
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   1020
      Width           =   315
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   435
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   300
   End
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mRedVal As Long
Public mGreenVal As Long
Public mBlueVal As Long

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
    Call Balancing(mdiImgProcess.ActiveForm.Tag)
    Screen.MousePointer = vbDefault
    ProgressBar.Value = 0
End Sub

Private Sub sldBlue_Change()
    mBlueVal = sldBlue.Value
End Sub

Private Sub sldBlue_Scroll()
    mBlueVal = sldBlue.Value
End Sub

Private Sub sldGreen_Change()
    mGreenVal = sldGreen.Value
End Sub

Private Sub sldGreen_Scroll()
    mGreenVal = sldGreen.Value
End Sub

Private Sub sldRed_Change()
    mRedVal = sldRed.Value
End Sub

Private Sub sldRed_Scroll()
    mRedVal = sldRed.Value
End Sub
