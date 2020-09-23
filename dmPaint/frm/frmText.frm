VERSION 5.00
Begin VB.Form frmText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboStyle 
      Height          =   315
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   255
      Width           =   1185
   End
   Begin VB.PictureBox pPreview 
      AutoRedraw      =   -1  'True
      Height          =   1140
      Left            =   4275
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   8
      Top             =   960
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   1845
      TabIndex        =   7
      Top             =   1665
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3000
      TabIndex        =   6
      Top             =   1665
      Width           =   1050
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Left            =   210
      TabIndex        =   5
      Top             =   1110
      Width           =   3825
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   3210
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   255
      Width           =   615
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   645
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   255
      Width           =   2100
   End
   Begin VB.Label lblStyle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Style:"
      Height          =   195
      Left            =   3900
      TabIndex        =   10
      Top             =   285
      Width           =   390
   End
   Begin VB.Label lblPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      Height          =   195
      Left            =   4260
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      Height          =   195
      Left            =   210
      TabIndex        =   4
      Top             =   855
      Width           =   360
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      Height          =   195
      Left            =   2805
      TabIndex        =   2
      Top             =   285
      Width           =   345
   End
   Begin VB.Label lblFont 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font:"
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   285
      Width           =   360
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function FindInCombo(Cbo As ComboBox, ByVal StrFind As String) As Integer
Dim X As Integer
Dim idx As Integer
    idx = 0
    
    If Len(StrFind) = 0 Then
        Exit Function
    Else
        For X = 0 To Cbo.ListCount
            If LCase(StrFind) = LCase(Cbo.List(X)) Then
                idx = X
                Exit For
            End If
        Next X
    End If
    
    FindInCombo = idx
End Function

Private Sub cboFont_Click()
On Error Resume Next
Dim sText As String

    sText = "Abc"
    
    With pPreview
        .Cls
        .FontSize = Val(cboSize.Text)
        .FontName = cboFont.Text
        .FontBold = False
        .FontItalic = False
        'Set font styles
        Select Case cboStyle.ListIndex
            Case 1
                .FontBold = True
            Case 2
                .FontItalic = True
            Case 3
                .FontBold = True
                .FontItalic = True
        End Select
        'Center Preview
        .CurrentX = (.ScaleWidth - .TextWidth(sText)) \ 2
        .CurrentY = (.ScaleHeight - .TextHeight(sText)) \ 2
        'Display preview
        pPreview.Print sText
        
    End With
    
End Sub

Private Sub cboSize_Click()
    Call cboFont_Click
End Sub

Private Sub cboStyle_Click()
    Call cboFont_Click
End Sub

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmText
End Sub

Private Sub cmdOK_Click()
    ButtonPress = vbOK
    'Fill in Text Type
    mTextProp.ffName = cboFont.Text
    mTextProp.fSize = Val(cboSize.Text)
    mTextProp.fText = txtText.Text
    mTextProp.fStyle = cboStyle.ListIndex
    Unload frmText
End Sub

Private Sub Form_Activate()
    txtText.SetFocus
End Sub

Private Sub Form_Load()
Dim Cnt As Integer
    
    Set frmText.Icon = Nothing
    'Add fonts
    For Cnt = 0 To (Screen.FontCount - 1)
        cboFont.AddItem Screen.Fonts(Cnt)
    Next Cnt
    
    'Add font sizes
    For Cnt = 8 To 72
        cboSize.AddItem Cnt
    Next Cnt
    
    'Add FontStyles
    cboStyle.AddItem "None"
    cboStyle.AddItem "Bold"
    cboStyle.AddItem "Italic"
    cboStyle.AddItem "Bold/Italic"
    
    txtText.Text = mTextProp.fText
    
    'Set Index's
    cboFont.ListIndex = FindInCombo(cboFont, mTextProp.ffName)
    cboSize.ListIndex = FindInCombo(cboSize, mTextProp.fSize)
    cboStyle.ListIndex = mTextProp.fStyle
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmText = Nothing
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call cmdOK_Click
    End If
End Sub
