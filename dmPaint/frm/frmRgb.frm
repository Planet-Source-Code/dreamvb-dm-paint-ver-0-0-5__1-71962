VERSION 5.00
Begin VB.Form frmRgb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjust Red/Blue/Green"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3945
      TabIndex        =   11
      Top             =   1440
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   2775
      TabIndex        =   10
      Top             =   1440
      Width           =   1050
   End
   Begin VB.PictureBox pView 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   3765
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   150
      Width           =   1200
   End
   Begin VB.HScrollBar sbRgb 
      Height          =   255
      Index           =   2
      Left            =   810
      Max             =   255
      Min             =   -255
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1005
      Width           =   2520
   End
   Begin VB.HScrollBar sbRgb 
      Height          =   255
      Index           =   1
      Left            =   810
      Max             =   255
      Min             =   -255
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   585
      Width           =   2520
   End
   Begin VB.HScrollBar sbRgb 
      Height          =   255
      Index           =   0
      Left            =   810
      Max             =   255
      Min             =   -255
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   225
      Width           =   2520
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   2
      Left            =   3405
      TabIndex        =   8
      Top             =   1020
      Width           =   90
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   1
      Left            =   3405
      TabIndex        =   7
      Top             =   615
      Width           =   90
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   3405
      TabIndex        =   6
      Top             =   270
      Width           =   90
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blue:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1035
      Width           =   360
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   585
      Width           =   480
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Red:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   345
   End
End
Attribute VB_Name = "frmRgb"
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
    Unload frmRgb
End Sub

Private Sub cmdOK_Click()
    'Fill in rgb type
    With TmpRgb
        .R = sbRgb(0).Value
        .G = sbRgb(1).Value
        .B = sbRgb(2).Value
    End With
    
    ButtonPress = vbOK
    Unload frmRgb
End Sub

Private Sub Form_Load()
    Set frmRgb.Icon = Nothing
    Call MakePreview
    pView.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRgb = Nothing
End Sub

Private Sub sbRgb_Change(Index As Integer)
    
    Call MakePreview
    
    lblValue(Index).Caption = sbRgb(Index).Value
    Call AdjustRGB(pView, sbRgb(0), sbRgb(1), sbRgb(2))
End Sub

Private Sub sbRgb_Scroll(Index As Integer)
    Call sbRgb_Change(Index)
End Sub
