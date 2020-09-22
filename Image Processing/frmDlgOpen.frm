VERSION 5.00
Begin VB.Form frmDlgOpen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open Picture"
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.FileListBox file 
      Height          =   3600
      Left            =   2700
      TabIndex        =   2
      Top             =   60
      Width           =   2235
   End
   Begin VB.DirListBox dir 
      Height          =   3240
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   2535
   End
   Begin VB.DriveListBox drv 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2535
   End
End
Attribute VB_Name = "frmDlgOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' I guess you should have known what are the meaning of these codes
'

Private Sub cmdCancel_Click()
    fpath = LoadPicture("")
    iCancel = True
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    fpath = file.Path & "\" & file.filename
    currDir = dir.Path
    Unload Me
End Sub

Private Sub dir_Change()
    file.Path = dir.Path
    file.Refresh
    cmdOpen.Enabled = False
End Sub

Private Sub drv_Change()
    dir.Path = drv.drive
    file.Path = dir.Path
    dir.Refresh
    file.Refresh
    cmdOpen.Enabled = False
End Sub

Private Sub file_Click()
    cmdOpen.Enabled = True
End Sub

Private Sub Form_Load()
    dir.Path = currDir
    dir.Refresh
    file.Refresh
    file.Pattern = "*.bmp;*.jpg"
    iCancel = False
End Sub
