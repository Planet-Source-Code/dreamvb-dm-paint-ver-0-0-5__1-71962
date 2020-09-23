VERSION 5.00
Begin VB.Form frmBright 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Brightness"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar sBar 
      Height          =   990
      Left            =   1395
      Max             =   255
      Min             =   -255
      TabIndex        =   4
      Top             =   375
      Width           =   360
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   1950
      TabIndex        =   3
      Top             =   840
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   1950
      TabIndex        =   2
      Top             =   390
      Width           =   1050
   End
   Begin VB.PictureBox pView 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   135
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   1
      Top             =   390
      Width           =   1200
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   1425
      TabIndex        =   5
      Top             =   1410
      Width           =   90
   End
   Begin VB.Label lblBrightness 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brightness"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmBright"
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

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmBright
End Sub

Private Sub CmdOK_Click()
    ButtonPress = vbYes
    'Store brightness value
    mBrightness = sBar.Value
    Unload frmBright
End Sub

Private Sub Form_Load()
    Set frmBright.Icon = Nothing
    Call MakePreview
    pView.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBright = Nothing
End Sub

Private Sub sBar_Change()
    Call MakePreview
    lblValue.Caption = sBar.Value
    'Change brightness
    Call LightenDarken(pView, sBar.Value)
End Sub

Private Sub sBar_Scroll()
    Call sBar_Change
End Sub
