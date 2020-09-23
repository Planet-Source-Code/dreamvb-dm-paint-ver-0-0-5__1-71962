VERSION 5.00
Begin VB.Form frmScale 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scale Image"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   300
      TabIndex        =   4
      Top             =   1200
      Width           =   1050
   End
   Begin VB.TextBox txtHeight 
      Height          =   350
      Left            =   1320
      TabIndex        =   3
      Top             =   630
      Width           =   915
   End
   Begin VB.TextBox txtWidth 
      Height          =   350
      Left            =   1320
      TabIndex        =   2
      Top             =   195
      Width           =   915
   End
   Begin VB.Label lblheight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image Height:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Width           =   990
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image Width:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "frmScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmScale
End Sub

Private Sub cmdOK_Click()
    'Set image width and height values
    mImage.iWidth = Val(txtWidth.Text)
    mImage.iHeight = Val(txtHeight.Text)
    
    If (mImage.iHeight <= 0) Or (mImage.iWidth <= 0) Then
        MsgBox "Bitmaps must be greater than zero pixels.", vbExclamation, frmNew.Caption
        Exit Sub
    End If
    
    ButtonPress = vbOK
    Unload frmScale
    
End Sub

Private Sub Form_Load()
    Set frmScale.Icon = Nothing
    'Set width and height
    txtWidth.Text = mImage.iWidth
    txtHeight.Text = mImage.iHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmScale = Nothing
End Sub
