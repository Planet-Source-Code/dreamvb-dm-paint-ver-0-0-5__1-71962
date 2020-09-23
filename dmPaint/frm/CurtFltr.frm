VERSION 5.00
Begin VB.Form frmCustFltr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Filter"
   ClientHeight    =   3030
   ClientLeft      =   3930
   ClientTop       =   5625
   ClientWidth     =   4140
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   33
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Process Now"
      Height          =   375
      Left            =   2280
      TabIndex        =   32
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   390
      Left            =   1440
      TabIndex        =   29
      Text            =   "127"
      Top             =   2505
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Text            =   "1"
      Top             =   2505
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   24
      Left            =   1680
      TabIndex        =   24
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   23
      Left            =   1320
      TabIndex        =   23
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   22
      Left            =   960
      TabIndex        =   22
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   21
      Left            =   600
      TabIndex        =   21
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   20
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   19
      Left            =   1680
      TabIndex        =   19
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   18
      Left            =   1320
      TabIndex        =   18
      Text            =   "-1"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   17
      Left            =   960
      TabIndex        =   17
      Text            =   "0"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   16
      Left            =   600
      TabIndex        =   16
      Text            =   "-1"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   15
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   14
      Left            =   1680
      TabIndex        =   14
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   13
      Left            =   1320
      TabIndex        =   13
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   12
      Left            =   960
      TabIndex        =   12
      Text            =   "4"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   11
      Left            =   600
      TabIndex        =   11
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   9
      Left            =   1680
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Text            =   "-1"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   7
      Left            =   960
      TabIndex        =   7
      Text            =   "0"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Text            =   "-1"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox SSPanel1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2280
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   25
      Top             =   225
      Width           =   1575
      Begin VB.OptionButton Option2 
         Caption         =   "5 X 5"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3 X 3 "
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Size"
         Height          =   195
         Left            =   375
         TabIndex        =   34
         Top             =   120
         Width           =   705
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bias"
      Height          =   195
      Left            =   1440
      TabIndex        =   31
      Top             =   2235
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Divide"
      Height          =   195
      Left            =   225
      TabIndex        =   30
      Top             =   2235
      Width           =   450
   End
End
Attribute VB_Name = "frmCustFltr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    frmCustFltr.Hide
End Sub

Private Sub CmdOK_Click()
Dim i As Integer
Dim j As Integer
    
    For i = 0 To 4
        For j = 0 To 4
            CustomFilter(i, j) = Val(Text1(i * 5 + j).Text)
        Next j
    Next i

    FilterNorm = Val(Text2.Text)
    FilterBias = Val(Text3.Text)
    ButtonPress = vbOK
    frmCustFltr.Hide
    
End Sub

Private Sub Form_Load()
    Set frmCustFltr.Icon = Nothing
    Call Option1_Click
End Sub

Private Sub Option1_Click()
Dim i As Integer

    For i = 0 To 4
        Text1(i).Visible = False
        Text1(i + 20).Visible = False
    Next i

    For i = 1 To 3
        Text1(i * 5).Visible = False
        Text1(i * 5 + 4).Visible = False
    Next i
    
    FliterSize = 1
End Sub

Private Sub Option2_Click()
Dim i As Integer

    For i = 0 To 4
        Text1(i).Visible = True
        Text1(i + 20).Visible = True
    Next i
    
    For i = 1 To 3
        Text1(i * 5).Visible = True
        Text1(i * 5 + 4).Visible = True
    Next i
    
    FliterSize = 2
End Sub

