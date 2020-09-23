VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Image"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3030
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pColor 
      Height          =   285
      Left            =   1935
      ScaleHeight     =   225
      ScaleWidth      =   270
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Select"
      Top             =   1395
      Width           =   330
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   165
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1980
      Width           =   1755
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   15
      TabIndex        =   8
      Top             =   2415
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   1560
      TabIndex        =   7
      Top             =   2580
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   2640
      TabIndex        =   6
      Top             =   2580
      Width           =   1050
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Left            =   165
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1380
      Width           =   1755
   End
   Begin VB.TextBox txtHeight 
      Height          =   315
      Left            =   795
      TabIndex        =   2
      Text            =   "300"
      Top             =   705
      Width           =   1020
   End
   Begin VB.TextBox txtWidth 
      Height          =   315
      Left            =   810
      TabIndex        =   1
      Text            =   "300"
      Top             =   255
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Preset size:"
      Height          =   195
      Left            =   165
      TabIndex        =   11
      Top             =   1785
      Width           =   810
   End
   Begin VB.Label lblPixel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixels"
      Height          =   195
      Index           =   1
      Left            =   1950
      TabIndex        =   10
      Top             =   795
      Width           =   405
   End
   Begin VB.Label lblPixel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixels"
      Height          =   195
      Index           =   0
      Left            =   1950
      TabIndex        =   9
      Top             =   315
      Width           =   405
   End
   Begin VB.Label lblImagebk 
      AutoSize        =   -1  'True
      Caption         =   "Image background color"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   1140
      Width           =   1725
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      Height          =   195
      Left            =   165
      TabIndex        =   3
      Top             =   765
      Width           =   510
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   330
      Width           =   465
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboColor_Click()
    Select Case cboColor.ListIndex
        Case 0: pColor.BackColor = vbWhite
        Case 1: pColor.BackColor = vbBlack
        Case 2: pColor.BackColor = vbRed
        Case 3: pColor.BackColor = vbGreen
        Case 4: pColor.BackColor = vbBlue
        Case 5: pColor.BackColor = frmmain.dColorPallete1.Color1
        Case 6: pColor.BackColor = frmmain.dColorPallete1.Color2
    End Select
End Sub

Private Sub cboSize_Click()
    Select Case cboSize.ListIndex
        Case 0
            txtWidth.Text = 100
            txtHeight.Text = 100
        Case 1
            txtWidth.Text = 128
            txtHeight.Text = 128
        Case 2
            txtWidth.Text = 200
            txtHeight.Text = 200
        Case 3
            txtWidth.Text = 300
            txtHeight.Text = 300
        Case 4
            txtWidth.Text = 800
            txtHeight.Text = 600
    End Select
End Sub

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmNew
End Sub

Private Sub CmdOK_Click()

    With mImage
        'Set image with and height
        .iHeight = Val(txtHeight.Text)
        .iWidth = Val(txtWidth.Text)
        If (.iHeight <= 0) Or (.iWidth <= 0) Then
            MsgBox "Bitmaps must be greater than zero pixels.", vbExclamation, frmNew.Caption
            Exit Sub
        End If
        'Set image background color
        .iBackColor = pColor.BackColor
    End With
    
    ButtonPress = vbOK
    Unload frmNew
    
End Sub

Private Sub Form_Load()
    Set frmNew.Icon = Nothing
    'Add some backcolors
    cboColor.AddItem "White"
    cboColor.AddItem "Black"
    cboColor.AddItem "Red"
    cboColor.AddItem "Green"
    cboColor.AddItem "Blue"
    cboColor.AddItem "Foreground"
    cboColor.AddItem "Background"
    'Add some preset sizes
    cboSize.AddItem "100x100"
    cboSize.AddItem "128x128"
    cboSize.AddItem "200x200"
    cboSize.AddItem "300x300"
    cboSize.AddItem "800x600"
    cboSize.ListIndex = 3
    cboColor.ListIndex = 0
End Sub

Private Sub pColor_Click()
On Error GoTo ColErr:
    'Update pColor with dialog selected color
    CD1.CancelError = True
    CD1.ShowColor
    pColor.BackColor = CD1.Color
    Exit Sub
    'Error flag
ColErr:
    If (Err.Number = cdlCancel) Then Err.Clear
    
End Sub
