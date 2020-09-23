VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStroke 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stroke Image"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD2 
      Left            =   2025
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   350
      Left            =   300
      TabIndex        =   5
      Top             =   1215
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   1050
   End
   Begin VB.TextBox txtWidth 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Text            =   "1"
      Top             =   570
      Width           =   735
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   1200
      ScaleHeight     =   255
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   210
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Stroke Width:"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   630
      Width           =   975
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      Caption         =   "Stroke Color:"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   915
   End
End
Attribute VB_Name = "frmStroke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmStroke
End Sub

Private Sub cmdOK_Click()
    StrokeC = pColor.BackColor
    StrokeW = Val(txtWidth.Text)
    
    ButtonPress = vbOK
    Unload frmStroke
End Sub

Private Sub Form_Load()
    Set frmStroke.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmStroke = Nothing
End Sub

Private Sub pColor_Click()
On Error GoTo CanErr:

    CD2.CancelError = True
    'Show color dialog box
    CD2.ShowColor
    'Update picturebox with color
    pColor.BackColor = CD2.Color
    Exit Sub
CanErr:
    If (Err = cdlCancel) Then
        Err.Clear
    End If
End Sub

Private Sub txtWidth_LostFocus()
    If Val(txtWidth.Text) = 0 Then
        txtWidth.Text = "1"
    End If
End Sub
