VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   2070
      TabIndex        =   2
      Top             =   1785
      Width           =   1050
   End
   Begin VB.PictureBox pTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   0
      Width           =   3285
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Free paint editor for windows"
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   2025
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   60
         Picture         =   "frmabout.frx":0000
         Top             =   75
         Width           =   2610
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "By DreamVB"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   210
      Left            =   90
      TabIndex        =   3
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "DM Paint 0.0.5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   1245
      Width           =   1380
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload frmabout
End Sub

Private Sub Form_Load()
    Set frmabout.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmabout = Nothing
End Sub
