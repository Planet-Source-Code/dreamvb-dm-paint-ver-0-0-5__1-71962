VERSION 5.00
Begin VB.Form frmRgbSwap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Swap RGB Channels"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   1500
      TabIndex        =   7
      Top             =   2040
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   2715
      TabIndex        =   6
      Top             =   2025
      Width           =   1050
   End
   Begin VB.PictureBox pView 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   2580
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   5
      Top             =   330
      Width           =   1200
   End
   Begin VB.OptionButton OptRGB 
      Caption         =   "RGB ==> BGR"
      Height          =   240
      Index           =   4
      Left            =   195
      TabIndex        =   4
      Top             =   1695
      Width           =   2010
   End
   Begin VB.OptionButton OptRGB 
      Caption         =   "RGB ==> BRG "
      Height          =   240
      Index           =   3
      Left            =   195
      TabIndex        =   3
      Top             =   1335
      Width           =   2010
   End
   Begin VB.OptionButton OptRGB 
      Caption         =   "RGB ==> GBR"
      Height          =   240
      Index           =   2
      Left            =   195
      TabIndex        =   2
      Top             =   975
      Width           =   2010
   End
   Begin VB.OptionButton OptRGB 
      Caption         =   "RGB ==> GRB"
      Height          =   240
      Index           =   1
      Left            =   195
      TabIndex        =   1
      Top             =   615
      Width           =   2010
   End
   Begin VB.OptionButton OptRGB 
      Caption         =   "RGB ==> RBG"
      Height          =   240
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   285
      Width           =   2010
   End
End
Attribute VB_Name = "frmRgbSwap"
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
    Unload frmRgbSwap
End Sub

Private Sub cmdOK_Click()
    ButtonPress = vbOK
    Unload frmRgbSwap
End Sub

Private Sub Form_Load()
    Set frmRgbSwap.Icon = Nothing
    'Display preview
    Call MakePreview
    Call OptRGB_Click(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRgbSwap = Nothing
End Sub

Private Sub OptRGB_Click(Index As Integer)
    mRGBSwapIndex = Index
    'Display preview
    Call MakePreview
    'Show rgb swap results
    Call SwapRGB(pView, mRGBSwapIndex)
End Sub
