VERSION 5.00
Begin VB.Form frmHighlight 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Highlight"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3975
      TabIndex        =   11
      Top             =   1455
      Width           =   1050
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   2835
      TabIndex        =   10
      Top             =   1455
      Width           =   1050
   End
   Begin VB.PictureBox pView 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   3795
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   6
      Top             =   120
      Width           =   1200
   End
   Begin VB.HScrollBar hRGB 
      Height          =   255
      Index           =   2
      Left            =   675
      Max             =   255
      TabIndex        =   5
      Top             =   885
      Width           =   2655
   End
   Begin VB.HScrollBar hRGB 
      Height          =   255
      Index           =   1
      Left            =   675
      Max             =   255
      TabIndex        =   4
      Top             =   510
      Width           =   2655
   End
   Begin VB.HScrollBar hRGB 
      Height          =   255
      Index           =   0
      Left            =   675
      Max             =   255
      TabIndex        =   3
      Top             =   150
      Width           =   2655
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   2
      Left            =   3390
      TabIndex        =   9
      Top             =   900
      Width           =   90
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   1
      Left            =   3390
      TabIndex        =   8
      Top             =   495
      Width           =   90
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   3390
      TabIndex        =   7
      Top             =   165
      Width           =   90
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   900
      Width           =   315
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   525
      Width           =   435
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   300
   End
End
Attribute VB_Name = "frmHighlight"
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
    Unload frmHighlight
End Sub

Private Sub CmdOK_Click()
    'Fill in rgb type
    With TmpRgb
        .R = hRGB(0).Value
        .G = hRGB(1).Value
        .B = hRGB(2).Value
    End With
    
    ButtonPress = vbOK
    Unload frmHighlight

End Sub

Private Sub Form_Load()
    Set frmHighlight.Icon = Nothing
    Call MakePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHighlight = Nothing
End Sub

Private Sub hRGB_Change(Index As Integer)
    lblVal(Index).Caption = hRGB(Index).Value
    'Display preview
    Call MakePreview
    'Do Highlight
    Call Highlight(pView, hRGB(0).Value, hRGB(1).Value, hRGB(2).Value)
End Sub

Private Sub hRGB_Scroll(Index As Integer)
    Call hRGB_Change(Index)
End Sub
