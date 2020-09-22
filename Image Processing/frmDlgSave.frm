VERSION 5.00
Begin VB.Form frmDlgSave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save Picture"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox file 
      Height          =   2235
      Left            =   3180
      Pattern         =   "*.bmp"
      TabIndex        =   8
      Top             =   60
      Width           =   3015
   End
   Begin VB.DirListBox dir 
      Height          =   1890
      Left            =   180
      TabIndex        =   7
      Top             =   420
      Width           =   2955
   End
   Begin VB.DriveListBox drive 
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   60
      Width           =   2955
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   3330
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3330
      Width           =   1215
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   900
      TabIndex        =   3
      Top             =   2820
      Width           =   4695
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2460
      Width           =   5235
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING : Existing file will be overwritten without prompt"
      Height          =   435
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   3300
      Width           =   3135
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".bmp"
      Height          =   195
      Index           =   2
      Left            =   5820
      TabIndex        =   9
      Top             =   2880
      Width           =   345
   End
   Begin VB.Line lin 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   8
      X2              =   412
      Y1              =   215
      Y2              =   215
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   0
      X1              =   8
      X2              =   412
      Y1              =   216
      Y2              =   216
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Directory"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   630
   End
End
Attribute VB_Name = "frmDlgSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    SavePicture mdiImgProcess.ActiveForm.pcbPicture.Image, txtDir.Text & txtFile.Text & ".bmp"
    Unload Me
End Sub

Private Sub dir_Change()
    file.Path = dir.Path
    file.Refresh
    txtDir.Text = dir.Path
    If Right(dir.Path, 1) <> "\" Then txtDir.Text = dir.Path & "\"
End Sub

Private Sub drive_Change()
    dir.Path = drive.drive
    dir.Refresh
    file.Path = dir.Path
    file.Refresh
End Sub

Private Sub file_Click()
    cmdSave.Enabled = True
    txtFile.Text = Left(file.filename, Len(file.filename) - 4)
End Sub

Private Sub Form_Load()
    dir.Path = currDir
    dir.Refresh
    file.Path = dir.Path
    file.Refresh
End Sub

Private Sub txtFile_Change()
    If Len(txtFile) <> 0 Then cmdSave.Enabled = True
    If Len(txtFile) = 0 Then cmdSave.Enabled = False
End Sub
