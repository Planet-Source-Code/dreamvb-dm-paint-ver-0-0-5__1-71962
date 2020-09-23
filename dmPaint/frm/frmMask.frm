VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMask 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Picture Mask"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   350
      Left            =   1590
      TabIndex        =   8
      Top             =   1335
      Width           =   1050
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   1590
      TabIndex        =   7
      Top             =   1785
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   1590
      TabIndex        =   6
      Top             =   915
      Width           =   1050
   End
   Begin VB.PictureBox pTmp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2820
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   5
      Top             =   1110
      Visible         =   0   'False
      Width           =   315
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3300
      Top             =   1125
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2820
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   4
      Top             =   1500
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "...."
      Height          =   345
      Left            =   3390
      TabIndex        =   3
      Top             =   390
      Width           =   480
   End
   Begin VB.TextBox txtFilename 
      Height          =   360
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   375
      Width           =   3150
   End
   Begin VB.PictureBox pView 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   135
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   0
      Top             =   915
      Width           =   1200
   End
   Begin VB.Label lblMask 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mask Filename:"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   165
      Width           =   1110
   End
End
Attribute VB_Name = "frmMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DoMask()
    'Resize temp picturebox
    pTmp.Width = pView.Width
    pTmp.Height = pView.Height
    
    'stretch mask onto pTmp
    StretchBlt pTmp.hDC, 0, 0, pView.Width, pView.Height, pMask.hDC, 0, 0, pMask.Width, pMask.Height, vbSrcCopy
    TransparentBlt pView.hDC, 0, 0, pView.Width, pView.Height, pTmp.hDC, 0, 0, pTmp.Width, pTmp.Height, RGB(255, 0, 255)
    pView.Refresh
    
    Set pTmp.Picture = Nothing
    Set pMask.Picture = Nothing
    
End Sub

Private Sub MakePreview()
    'Set pallete mode
    SetStretchBltMode pView.hDC, vbPaletteModeNone
    StretchBlt pView.hDC, 0, 0, 80, 80, frmmain.SrcDc.hDC, 0, 0, _
    frmmain.SrcDc.ScaleWidth, frmmain.SrcDc.ScaleHeight, vbSrcCopy
End Sub

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmMask
End Sub

Private Sub CmdOK_Click()
    'Store maskfilename
    mMaskFile = txtFilename.Text
    ButtonPress = vbOK
    Unload frmMask
End Sub

Private Sub cmdOpen_Click()
On Error GoTo OpenErr:

    With CD1
        .CancelError = True
        .DialogTitle = "Open"
        .ShowOpen
        txtFilename.Text = .FileName
        pMask.Picture = LoadPicture(.FileName)
        Call MakePreview
        Call DoMask
    End With
    
    Exit Sub
OpenErr:
    If Err.Number = cdlCancel Then
        Err.Clear
    End If
End Sub

Private Sub cmdReset_Click()
    Call MakePreview
    pView.Refresh
End Sub

Private Sub Form_Load()
    Set frmMask.Icon = Nothing
    Call MakePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMask = Nothing
End Sub
