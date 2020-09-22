VERSION 5.00
Begin VB.Form frmPicture 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1680
   ScaleWidth      =   1710
   Begin VB.PictureBox pcbPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pWidth As Single
Dim pHeight As Single

Private Sub Form_Load()
    pcbPicture.Picture = LoadPicture(fpath) 'Load the picture into picture box
    
    hDCSour(indeks) = CreateCompatibleDC(pcbPicture.hdc)    'Create a space in memory for source picture
    hBMPSour(indeks) = CreateCompatibleBitmap(pcbPicture.hdc, pcbPicture.ScaleWidth, pcbPicture.ScaleHeight)    'Create bitmap structures in memory
    hDCDest(indeks) = CreateCompatibleDC(pcbPicture.hdc)    'Create a space in memory for destination picture
    hBMPDest(indeks) = CreateCompatibleBitmap(pcbPicture.hdc, pcbPicture.ScaleWidth, pcbPicture.ScaleHeight)    'Create bitmap structures in memory
    SelectObject hDCSour(indeks), hBMPSour(indeks)
    SelectObject hDCDest(indeks), hBMPDest(indeks)  'select those spaces and bitmaps
    
    BitBlt hDCSour(indeks), 0, 0, pcbPicture.ScaleWidth, pcbPicture.ScaleHeight, pcbPicture.hdc, 0, 0, vbSrcCopy
        'Blit the picture inside picture box into memory

    Me.Width = pcbPicture.Width + 125
    Me.Height = pcbPicture.Height + 405
    pcbPicture.Left = (Me.ScaleWidth - pcbPicture.Width) / 2
    pcbPicture.Top = (Me.ScaleHeight - pcbPicture.Height) / 2
        'Adjust the picture box
End Sub

Private Sub Form_Resize()
    pcbPicture.Left = (Me.ScaleWidth - pcbPicture.Width) / 2
    pcbPicture.Top = (Me.ScaleHeight - pcbPicture.Height) / 2
        'Adjust the picture box
End Sub
