VERSION 5.00
Begin VB.Form frmGrade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gradient"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3300
      TabIndex        =   7
      Top             =   705
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   3300
      TabIndex        =   6
      Top             =   285
      Width           =   1050
   End
   Begin VB.ComboBox cboAlpha 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1155
      Width           =   1635
   End
   Begin VB.ComboBox cboDir 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   390
      Width           =   1635
   End
   Begin VB.PictureBox SrcPic 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   1770
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   1
      Top             =   3300
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox pView 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   1995
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   0
      Top             =   300
      Width           =   1200
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alpha Level:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   930
      Width           =   885
   End
   Begin VB.Label lbldir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direction:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   165
      Width           =   675
   End
End
Attribute VB_Name = "frmGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MakePreview()
    'Set pallete mode
    SetStretchBltMode pView.hDC, vbPaletteModeNone
    StretchBlt pView.hDC, 0, 0, 80, 80, frmmain.SrcDc.hDC, 0, 0, _
    frmmain.SrcDc.ScaleWidth, frmmain.SrcDc.ScaleHeight, vbSrcCopy
End Sub

Private Sub cboAlpha_Click()
    Call cboDir_Click
End Sub

Private Sub cboDir_Click()
    'Clear pic box
    Call pView.Cls
    'Display preview
    Call MakePreview
    'Grident picture
    Call GradientAlpha(pView, SrcPic, Val(cboAlpha.Text), frmmain.dColorPallete1.Color1, _
    frmmain.dColorPallete1.Color2, cboDir.ListIndex)
End Sub

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmGrade
End Sub

Private Sub CmdOK_Click()
    ButtonPress = vbOK
    'Set alpha level
    mAlphaLevel = Val(cboAlpha.Text)
    'Set alpha dir
    AlphaDir = cboDir.ListIndex
    Unload frmGrade
End Sub

Private Sub Form_Load()
Dim Cnt As Integer

    Set frmGrade.Icon = Nothing
    'Add gradient direction
    cboDir.AddItem "Left To Right"
    cboDir.AddItem "Top To Bottom"
    Call MakePreview
    
    'Add alpha levels
    For Cnt = 0 To 255
        cboAlpha.AddItem Cnt
    Next Cnt
    'Set indexs
    cboDir.ListIndex = AlphaDir
    cboAlpha.ListIndex = mAlphaLevel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGrade = Nothing
End Sub
