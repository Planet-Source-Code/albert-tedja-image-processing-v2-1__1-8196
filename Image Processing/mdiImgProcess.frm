VERSION 5.00
Begin VB.MDIForm mdiImgProcess 
   BackColor       =   &H8000000C&
   Caption         =   "Image Processing by Albert Nicholas"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8310
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrCheck 
      Interval        =   100
      Left            =   840
      Top             =   2220
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Picture..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Picture..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAdjust 
      Caption         =   "&Adjust"
      Begin VB.Menu mnuAdjustBright 
         Caption         =   "Brightness..."
      End
      Begin VB.Menu mnuAdjustCB 
         Caption         =   "Color Balance..."
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "F&ilters"
      Begin VB.Menu mnuFilterDiff 
         Caption         =   "Diffuse"
      End
      Begin VB.Menu mnuFilterEmb 
         Caption         =   "Emboss"
      End
      Begin VB.Menu mnuFilterGS 
         Caption         =   "Grayscale"
      End
      Begin VB.Menu mnuFilterInvert 
         Caption         =   "Invert Color"
      End
      Begin VB.Menu mnuFilterLE 
         Caption         =   "Lighting Effect"
      End
      Begin VB.Menu mnuFilterSharp 
         Caption         =   "Sharpen"
      End
      Begin VB.Menu mnuFilterSoft 
         Caption         =   "Soften"
      End
      Begin VB.Menu mnuFilterSolar 
         Caption         =   "Solarize"
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Window"
      Begin VB.Menu mnuWinCas 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWinHor 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuWinVer 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuWinArr 
         Caption         =   "Arrange Icons"
      End
   End
End
Attribute VB_Name = "mdiImgProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================================
' Image Processing
' Author : Albert Nicholas
' Email  : nicho_tedja@yahoo.com
' ===================================
' The last project version takes a very long time to process an image
' Now, by using Windows API, this project enables you to process an image
' in a much shorter time.
' I apologize for those who have waited long for this projects
' Enclosed:
' sample01.jpg
' sample02.jpg
' just for sample pictures, in case you do not own any :)

Private Sub MDIForm_Load()
    indeks = 0
    currDir = "C:\My Documents"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim i As Integer
    For i = 0 To indeks
        DeleteObject hBMPSour(i)
        DeleteDC hDCSour(i)
        DeleteObject hBMPDest(i)
        DeleteDC hDCDest(i)
            'THESE ARE IMPORTANT THINGS TO DO
            'Destroy all spaces and bitmaps to clean up memory
    Next i
    End
End Sub

Private Sub mnuAdjustBright_Click()
    Load frmBright
    frmBright.Show 1
End Sub

Private Sub mnuAdjustCB_Click()
    Load frmBalance
    frmBalance.Show 1
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileOpen_Click()
    Load frmDlgOpen
    frmDlgOpen.Show 1
    If iCancel = True Then Exit Sub
    picforms(indeks).Show
    picforms(indeks).Caption = fpath
    picforms(indeks).Tag = indeks
    indeks = indeks + 1
End Sub

Private Sub mnuFileSave_Click()
    Load frmDlgSave
    frmDlgSave.Show 1
End Sub

Private Sub mnuFilterDiff_Click()
    Screen.MousePointer = vbHourglass
    Call Diffusing(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilterEmb_Click()
    Screen.MousePointer = vbHourglass
    Call Embossing(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilterLE_Click()
    Screen.MousePointer = vbHourglass
    Call Lighting(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilterSC_Click()
    Screen.MousePointer = vbHourglass
    Call Switching(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilterDark_Click()
    Screen.MousePointer = vbHourglass
    Call Darken(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilterGS_Click()
    Screen.MousePointer = vbHourglass
    Call Grayscaling(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilterInvert_Click()
    Screen.MousePointer = vbHourglass
    Call Inverting(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilterLight_Click()
    Screen.MousePointer = vbHourglass
    Call Lighten(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilterSharp_Click()
    Screen.MousePointer = vbHourglass
    Call Sharpening(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilterSoft_Click()
    Screen.MousePointer = vbHourglass
    Call Softening(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilterSolar_Click()
    Screen.MousePointer = vbHourglass
    Call Solarizing(ActiveForm.Tag)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuWinArr_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWinCas_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWinHor_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWinVer_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub tmrCheck_Timer()
    If Me.ActiveForm Is Nothing Then
        mnuFileSave.Enabled = False
        mnuFilter.Enabled = False
        mnuAdjust.Enabled = False
    Else
        mnuFileSave.Enabled = True
        mnuFilter.Enabled = True
        mnuAdjust.Enabled = True
    End If
End Sub
