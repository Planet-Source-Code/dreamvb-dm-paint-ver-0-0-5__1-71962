VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "DreamVB's Paint"
   ClientHeight    =   10680
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   712
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   970
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2085
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   65
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox SrcBrush1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   2085
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   45
      Top             =   5835
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox SrcBrush2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   2085
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   44
      Top             =   5355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pSrc2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2115
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   26
      Top             =   4995
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8265
      Top             =   780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pTools 
      Height          =   9735
      Left            =   15
      ScaleHeight     =   9675
      ScaleWidth      =   1890
      TabIndex        =   7
      Top             =   630
      Width           =   1950
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   23
         Left            =   0
         TabIndex        =   67
         Tag             =   "0"
         Top             =   8880
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Arrow"
         Picture         =   "frmmain.frx":0000
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   22
         Left            =   0
         TabIndex        =   66
         Tag             =   "0"
         Top             =   6540
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Triangle"
         Picture         =   "frmmain.frx":0352
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   21
         Left            =   0
         TabIndex        =   61
         Tag             =   "0"
         Top             =   8490
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Text"
         Picture         =   "frmmain.frx":06A4
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   57
         Tag             =   "0"
         Top             =   315
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Select"
         Picture         =   "frmmain.frx":09F6
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   19
         Left            =   0
         TabIndex        =   53
         Tag             =   "1"
         Top             =   8100
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Clone"
         Picture         =   "frmmain.frx":0B08
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   52
         Tag             =   "0"
         Top             =   1485
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Star"
         Picture         =   "frmmain.frx":0C1A
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   16
         Left            =   0
         TabIndex        =   43
         Tag             =   "7"
         Top             =   6930
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Brush"
         Picture         =   "frmmain.frx":0F6C
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   15
         Left            =   0
         TabIndex        =   29
         Tag             =   "0"
         Top             =   6150
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Polygone"
         Picture         =   "frmmain.frx":107E
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   20
         Left            =   0
         TabIndex        =   28
         Tag             =   "0"
         Top             =   9270
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Gradient"
         Picture         =   "frmmain.frx":13D0
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   18
         Left            =   0
         TabIndex        =   27
         Tag             =   "3"
         Top             =   7710
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Color Picker"
         Picture         =   "frmmain.frx":1722
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   17
         Left            =   0
         TabIndex        =   25
         Tag             =   "6"
         Top             =   7320
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Erase"
         Picture         =   "frmmain.frx":1A74
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   14
         Left            =   0
         TabIndex        =   24
         Tag             =   "0"
         Top             =   5760
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Curve"
         Picture         =   "frmmain.frx":1DC6
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   13
         Left            =   0
         TabIndex        =   23
         Tag             =   "2"
         Top             =   5370
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Spray Can"
         Picture         =   "frmmain.frx":2118
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   4
         Left            =   0
         TabIndex        =   14
         Tag             =   "0"
         Top             =   1860
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Box"
         Picture         =   "frmmain.frx":246A
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Tag             =   "0"
         Top             =   1095
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Line"
         Picture         =   "frmmain.frx":27BC
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Tag             =   "5"
         Top             =   705
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Pencil"
         Picture         =   "frmmain.frx":2B0E
      End
      Begin VB.PictureBox pTitlebar 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   1905
         TabIndex        =   10
         Top             =   0
         Width           =   1905
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Drawing Tools"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   345
            TabIndex        =   11
            Top             =   45
            Width           =   1230
         End
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   5
         Left            =   0
         TabIndex        =   15
         Tag             =   "0"
         Top             =   2250
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Filled Box"
         Picture         =   "frmmain.frx":2E60
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   6
         Left            =   0
         TabIndex        =   16
         Tag             =   "0"
         Top             =   2640
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Circle"
         Picture         =   "frmmain.frx":31B2
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   7
         Left            =   0
         TabIndex        =   17
         Tag             =   "0"
         Top             =   3030
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Filled Circle"
         Picture         =   "frmmain.frx":3504
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   8
         Left            =   0
         TabIndex        =   18
         Tag             =   "0"
         Top             =   3420
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ellipse"
         Picture         =   "frmmain.frx":3856
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   9
         Left            =   0
         TabIndex        =   19
         Tag             =   "0"
         Top             =   3810
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Filled Ellipse"
         Picture         =   "frmmain.frx":3BA8
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   10
         Left            =   0
         TabIndex        =   20
         Tag             =   "0"
         Top             =   4200
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "RoundRect"
         Picture         =   "frmmain.frx":3EFA
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   11
         Left            =   0
         TabIndex        =   21
         Tag             =   "0"
         Top             =   4590
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Filled RoundRect"
         Picture         =   "frmmain.frx":424C
      End
      Begin Project1.dFlatButton cmdtool 
         Height          =   390
         Index           =   12
         Left            =   0
         TabIndex        =   22
         Tag             =   "4"
         Top             =   4980
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Flood Fill"
         Picture         =   "frmmain.frx":459E
      End
   End
   Begin VB.PictureBox pHolder 
      BackColor       =   &H00808080&
      Height          =   4245
      Left            =   1995
      ScaleHeight     =   279
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   406
      TabIndex        =   3
      Top             =   630
      Width           =   6150
      Begin VB.CommandButton pSpacer 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4980
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   60
         Width           =   270
      End
      Begin VB.VScrollBar vBar 
         Height          =   450
         Left            =   4620
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   270
      End
      Begin VB.HScrollBar hBar 
         Height          =   270
         Left            =   0
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3285
         Width           =   480
      End
      Begin VB.PictureBox SrcDc 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4500
         Left            =   0
         MousePointer    =   2  'Cross
         ScaleHeight     =   300
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   6
         Top             =   15
         Width           =   4500
         Begin VB.Shape shpSelect 
            BorderStyle     =   3  'Dot
            Height          =   450
            Left            =   240
            Top             =   360
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Shape shpClone 
            Height          =   90
            Left            =   270
            Top             =   240
            Visible         =   0   'False
            Width           =   90
         End
      End
      Begin VB.Image pGrip 
         Height          =   90
         Index           =   2
         Left            =   4665
         MousePointer    =   7  'Size N S
         Top             =   1200
         Width           =   90
      End
      Begin VB.Image pGrip 
         Height          =   90
         Index           =   1
         Left            =   4665
         MousePointer    =   8  'Size NW SE
         Top             =   1050
         Width           =   90
      End
      Begin VB.Image pGrip 
         Height          =   90
         Index           =   0
         Left            =   4680
         MousePointer    =   9  'Size W E
         Picture         =   "frmmain.frx":48F0
         Top             =   900
         Width           =   90
      End
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   10410
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20029
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   14550
      TabIndex        =   0
      Top             =   0
      Width           =   14550
      Begin VB.PictureBox pTextStyle 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3900
         ScaleHeight     =   495
         ScaleWidth      =   3660
         TabIndex        =   62
         Top             =   75
         Visible         =   0   'False
         Width           =   3660
         Begin VB.ComboBox cboTextStyle 
            Height          =   315
            Left            =   945
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   60
            Width           =   1125
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text Style:"
            Height          =   195
            Left            =   30
            TabIndex        =   64
            Top             =   120
            Width           =   750
         End
      End
      Begin VB.PictureBox pGrad 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3900
         ScaleHeight     =   495
         ScaleWidth      =   3660
         TabIndex        =   58
         Top             =   75
         Visible         =   0   'False
         Width           =   3660
         Begin VB.ComboBox cboGrade 
            Height          =   315
            Left            =   765
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   60
            Width           =   1710
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direction:"
            Height          =   195
            Left            =   30
            TabIndex        =   60
            Top             =   120
            Width           =   675
         End
      End
      Begin VB.PictureBox pClone 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3900
         ScaleHeight     =   495
         ScaleWidth      =   3660
         TabIndex        =   54
         Top             =   75
         Visible         =   0   'False
         Width           =   3660
         Begin VB.ComboBox cboCloneSize 
            Height          =   315
            Left            =   855
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   60
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clone Size:"
            Height          =   195
            Left            =   30
            TabIndex        =   56
            Top             =   120
            Width           =   795
         End
      End
      Begin VB.PictureBox pBrush 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3900
         ScaleHeight     =   495
         ScaleWidth      =   3660
         TabIndex        =   46
         Top             =   75
         Visible         =   0   'False
         Width           =   3660
         Begin VB.ComboBox cboBrush 
            Height          =   315
            Left            =   525
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   60
            Width           =   1725
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brush:"
            Height          =   195
            Left            =   30
            TabIndex        =   48
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.PictureBox pErase 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3900
         ScaleHeight     =   495
         ScaleWidth      =   4230
         TabIndex        =   40
         Top             =   75
         Width           =   4230
         Begin VB.ComboBox cboEraseType 
            Height          =   315
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   60
            Width           =   1065
         End
         Begin VB.ComboBox cboWidth 
            Height          =   315
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Style:"
            Height          =   195
            Left            =   2040
            TabIndex        =   51
            Top             =   120
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DrawWdith:"
            Height          =   195
            Left            =   30
            TabIndex        =   42
            Top             =   120
            Width           =   840
         End
      End
      Begin VB.PictureBox pDrawWidth 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3900
         ScaleHeight     =   495
         ScaleWidth      =   5505
         TabIndex        =   35
         Top             =   75
         Visible         =   0   'False
         Width           =   5505
         Begin VB.ComboBox cboArrowSize 
            Height          =   315
            Left            =   4425
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cboStyle 
            Height          =   315
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   60
            Width           =   1065
         End
         Begin VB.ComboBox cboDW 
            Height          =   315
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label lblArrowSize 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Arrow Size:"
            Height          =   195
            Left            =   3600
            TabIndex        =   68
            Top             =   120
            Width           =   795
         End
         Begin VB.Label lblStyle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Style:"
            Height          =   195
            Left            =   2040
            TabIndex        =   38
            Top             =   120
            Width           =   390
         End
         Begin VB.Label lbldWidth 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DrawWidth:"
            Height          =   195
            Left            =   30
            TabIndex        =   37
            Top             =   120
            Width           =   840
         End
      End
      Begin VB.PictureBox pFillType 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3900
         ScaleHeight     =   495
         ScaleWidth      =   4650
         TabIndex        =   30
         Top             =   75
         Width           =   4650
         Begin VB.ComboBox cboPatten 
            Height          =   315
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   60
            Width           =   1950
         End
         Begin VB.ComboBox cboFillType 
            Height          =   315
            Left            =   705
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label lblPatten 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Patten:"
            Height          =   195
            Left            =   1860
            TabIndex        =   33
            Top             =   120
            Width           =   510
         End
         Begin VB.Label lblFillType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fill Type:"
            Height          =   195
            Left            =   30
            TabIndex        =   31
            Top             =   120
            Width           =   630
         End
      End
      Begin Project1.Line3D lnTop2 
         Height          =   30
         Left            =   0
         TabIndex        =   9
         Top             =   570
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   53
      End
      Begin Project1.Line3D lnTop1 
         Height          =   30
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   53
      End
      Begin Project1.dColorPallete dColorPallete1 
         Height          =   480
         Left            =   15
         TabIndex        =   1
         Top             =   60
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   847
      End
   End
   Begin VB.Image ImgPatten 
      Height          =   360
      Left            =   2085
      Top             =   6885
      Width           =   405
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuBlank6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select &All"
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Begin VB.Menu mnuFlip 
         Caption         =   "&Flip"
         Begin VB.Menu mnuFlipV 
            Caption         =   "&Vertical"
         End
         Begin VB.Menu mnuFlipHoz 
            Caption         =   "&Horizonal"
         End
         Begin VB.Menu mnuBoth 
            Caption         =   "&Both"
         End
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScale 
         Caption         =   "&Scale Image"
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "&Crop Image"
      End
      Begin VB.Menu mnuBlank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear Image"
      End
      Begin VB.Menu MnuDQue 
         Caption         =   "&Draw Opaque"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuAdjust0 
      Caption         =   "&Adjust"
      Begin VB.Menu mnuBright 
         Caption         =   "Brightness"
      End
      Begin VB.Menu mnuGamma 
         Caption         =   "Gramma"
      End
      Begin VB.Menu mnuContrast 
         Caption         =   "Contrast"
      End
      Begin VB.Menu mnuBlank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReplaceC 
         Caption         =   "Replace Color"
      End
      Begin VB.Menu mnuSwapRGB 
         Caption         =   "&Swap RGB Channels"
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "&Filter"
      Begin VB.Menu mnuEffect 
         Caption         =   "&Effects"
         Begin VB.Menu mnuMelt 
            Caption         =   "&Melt"
         End
         Begin VB.Menu mnuMirror 
            Caption         =   "&Mirror Left"
         End
         Begin VB.Menu mnuMRight 
            Caption         =   "Mirror Right"
         End
         Begin VB.Menu mnuGrid 
            Caption         =   "&Grid Generator"
         End
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "&Frame"
         Begin VB.Menu mnuStroke 
            Caption         =   "Stroke"
         End
         Begin VB.Menu mnuMask 
            Caption         =   "&Mask"
         End
      End
      Begin VB.Menu mnuNoise 
         Caption         =   "Noise"
         Begin VB.Menu mnuAddNoise 
            Caption         =   "&Add Noise"
         End
      End
      Begin VB.Menu mnuSoft 
         Caption         =   "&Soften"
         Begin VB.Menu mnuBlur 
            Caption         =   "&Blur"
         End
         Begin VB.Menu mnuBlureMore 
            Caption         =   "Blure More"
         End
      End
      Begin VB.Menu mnuSharp 
         Caption         =   "&Sharpen"
         Begin VB.Menu mnuSharpen 
            Caption         =   "Sharpen"
         End
         Begin VB.Menu mnuSharpMore 
            Caption         =   "Sharpen More"
         End
      End
      Begin VB.Menu mnuFlaten 
         Caption         =   "&Flatten"
         Begin VB.Menu mnuImboss 
            Caption         =   "&Imboss"
         End
         Begin VB.Menu mnuPlaster 
            Caption         =   "&Plaster"
         End
      End
      Begin VB.Menu mnuRender 
         Caption         =   "&Render"
         Begin VB.Menu mnuPixel 
            Caption         =   "&Pixel Effect"
         End
      End
      Begin VB.Menu mnuDiffuse 
         Caption         =   "&Diffuse"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "&Color"
         Begin VB.Menu mnuCBurn 
            Caption         =   "&Color Burn"
         End
         Begin VB.Menu mnuAdjust 
            Caption         =   "&Adjust Red/Green/Blue"
         End
         Begin VB.Menu mnuBlackWhite 
            Caption         =   "&BlackAndWhite"
         End
         Begin VB.Menu mnuHighlight 
            Caption         =   "&Highlight"
         End
         Begin VB.Menu mnuBlank5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGray 
            Caption         =   "&Grayscale"
         End
         Begin VB.Menu mnuInvert 
            Caption         =   "&Invert"
            Begin VB.Menu mnuAll 
               Caption         =   "All"
            End
            Begin VB.Menu mnuBlank7 
               Caption         =   "-"
            End
            Begin VB.Menu mnuRed1 
               Caption         =   "Red"
            End
            Begin VB.Menu mnuGreen1 
               Caption         =   "Green"
            End
            Begin VB.Menu mnuBlue1 
               Caption         =   "Blue"
            End
         End
         Begin VB.Menu mnuXorC 
            Caption         =   "XORColor"
         End
         Begin VB.Menu mnuRemove 
            Caption         =   "&Remove"
            Begin VB.Menu mnuRed 
               Caption         =   "Red"
            End
            Begin VB.Menu mnuGreen 
               Caption         =   "Green"
            End
            Begin VB.Menu mnuBlue 
               Caption         =   "Blue"
            End
         End
      End
      Begin VB.Menu mnuOther 
         Caption         =   "&Other"
         Begin VB.Menu mnuGrade 
            Caption         =   "Gradient"
         End
         Begin VB.Menu mnuSnow 
            Caption         =   "Snow"
         End
         Begin VB.Menu mnuCustom 
            Caption         =   "&Custom"
         End
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum DrawTools
    tSelect = 0
    tPencil = 1
    tline = 2
    tStar = 3
    tBox = 4
    tBoxF = 5
    tCircle = 6
    tCircleF = 7
    tEllipse = 8
    tEllipseF = 9
    tRoundRect = 10
    tRoundRectF = 11
    tFillCan = 12
    tSpray = 13
    tCurve = 14
    tPoly = 15
    tBrush = 16
    tErase = 17
    tColorPicker = 18
    tClone = 19
    tGradient1 = 20
    tText = 21
    tTriangle = 22
    tArrow = 23
End Enum

Private Type DrawPos
    StartX As Single
    StartY As Single
    OldX As Single
    OldY As Single
    X1 As Single
    Y1 As Single
End Type

Private mDrawPos As DrawPos
Private DrawTool As DrawTools
Private CanDraw As Boolean
Private IsDown As Boolean
Private CanPaste As Boolean
Private CopySelect As Boolean
Private mX As Integer

'Grip Variables
Private IsMoveing As Boolean
Private MoveX As Single
Private MoveY As Single
'Dialog Variables
Private dInitDir As String
Private Const dFilter1 = "Images Files (*.bmp;*.gif;*.jpg;*.png;*.tif)|*.bmp;*.gif;*.jpg;*.png;*.tif"
Private Const dFilter2 = "Bitmap Files(*.bmp)|*.bmp|PNG Files(*.png)|*.png|Gif Files(*.gif)|*.gif|JPEG Files(*.jpg)|*.jpg|"

Private Const FLOODFILLSURFACE As Long = 1

'Clone tool variables
Dim TmpX As Single
Dim TmpY As Single
Private IsCloneing As Boolean

Private Sub SetCursor(ByVal Index As Integer)
    If (Index = 0) Then
        SrcDc.MousePointer = vbCrosshair
    Else
        SrcDc.MousePointer = vbCustom
        SrcDc.MouseIcon = LoadResPicture(Index, vbResCursor)
    End If
    
End Sub

Private Sub DrawArrow(pBox As PictureBox, X1 As Single, Y1 As Single, X2 As Single, ByVal Y2 As Single, ArrowLength As Single)
Dim X As Single
Dim Y As Single
Dim Length As Single
    'This draws an arrow
    pBox.Line (X1, Y1)-(X2, Y2)
    X = X2 - X1
    Y = Y2 - Y1
    Length = Sqr(X * X + Y * Y)
    If (Length = 0) Then Length = 1
    X = (X / Length * ArrowLength)
    Y = (Y / Length * ArrowLength)
    'Draw the arrow head
    pBox.Line (X2, Y2)-Step(-X - Y, -Y + X)
    pBox.Line (X2, Y2)-Step(-X + Y, -Y - X)
End Sub

Private Sub DrawGrid(pBox As PictureBox, ByVal GridSize As Integer)
Dim X As Long
Dim Y As Long

    'This draws a grid on the picture

    For X = 0 To pBox.ScaleHeight Step GridSize
       pBox.Line (0, X)-(pBox.ScaleWidth - 1, X), dColorPallete1.Color1
    Next X
    
    For X = 0 To pBox.ScaleWidth Step GridSize
        pBox.Line (X, 0)-(X, pBox.ScaleHeight - 1), dColorPallete1.Color1
    Next X

End Sub

Private Sub ApplyImageMask(ByVal MaskFilename As String)
    'Load maskfilename
    pMask.Picture = LoadPicture(MaskFilename)
    
    SrcBrush2.Width = SrcDc.Width
    SrcBrush2.Height = SrcDc.Height
    
    SetStretchBltMode SrcBrush2.hdc, vbPaletteModeNone
    StretchBlt SrcBrush2.hdc, 0, 0, SrcDc.Width, SrcDc.Height, pMask.hdc, 0, 0, pMask.Width, pMask.Height, vbSrcCopy
    TransparentBlt SrcDc.hdc, 0, 0, SrcDc.Width, SrcDc.Height, SrcBrush2.hdc, 0, 0, SrcBrush2.Width, SrcBrush2.Height, RGB(255, 0, 255)
    SrcDc.Refresh
    
    Set pMask.Picture = Nothing
    Set SrcBrush2.Picture = Nothing
End Sub

Private Sub CutCopyImage(iCopy As Boolean, Offset As Integer, Optional CutImage As Boolean = True)
    If (iCopy) Then
        pSrc2.Width = (shpSelect.Width + Offset)
        pSrc2.Height = (shpSelect.Height + Offset)
        'Copy selection on to temp picturebox
        BitBlt pSrc2.hdc, 0, 0, pSrc2.Width, pSrc2.Height, SrcDc.hdc, shpSelect.Left, shpSelect.Top, vbSrcCopy
        pSrc2.Refresh
        'Copy image data to clipboard
        Call PictureBoxToClipboard(pSrc2)
        'Check if need to cut the image away
        If (CutImage) Then
            'Cut the selection
            SrcDc.Line (shpSelect.Left, shpSelect.Top)-(shpSelect.Width + shpSelect.Left, _
            shpSelect.Height + shpSelect.Top), SrcDc.BackColor, BF
            SrcDc.Refresh
        End If
    End If
End Sub

Private Sub DrawBrush()
Dim X As Long
Dim Y As Long
Dim Col As Long
Dim iRet As Long

    'Set the brush to be copyied
    SrcBrush2.Picture = SrcBrush1.Picture
    
    For X = 0 To SrcBrush2.ScaleWidth
        For Y = 0 To SrcBrush2.ScaleHeight
            Col = GetPixel(SrcBrush2.hdc, X, Y)
            If (Col = vbBlack) Then
                Col = SrcDc.ForeColor
                iRet = SetPixelV(SrcBrush2.hdc, X, Y, Col)
            End If
        Next Y
    Next X
    
    SrcBrush2.Refresh
    
End Sub

Private Sub LoadBrushs()
Dim xFile As String
    'This sub Loads in the brushs
    xFile = Dir(DataDir & "brushs\*.bmp")

    Do Until (xFile = "")
        cboBrush.AddItem xFile
        'Get next filename
        xFile = Dir()
        DoEvents
    Loop
    
    If (cboBrush.ListCount) Then
        cboBrush.ListIndex = 0
    End If
    
End Sub

Private Sub LoadPattens()
Dim xFile As String
    'This sub Loads in the pattens
    xFile = Dir(DataDir & "pattens\*.bmp")
    'Clear the combobox
    cboPatten.Clear
    Do Until (xFile = "")
        cboPatten.AddItem xFile
        'Get next filename
        xFile = Dir()
        DoEvents
    Loop
    
    If (cboPatten.ListCount) Then
        cboPatten.ListIndex = 0
    End If
    
End Sub

Private Sub LoadHatchs()
    'Clear combobox
    cboPatten.Clear
    'Add hatchs
    cboPatten.AddItem "HORIZONTAL"
    cboPatten.AddItem "VERTICAL"
    cboPatten.AddItem "FDIAGONAL"
    cboPatten.AddItem "BDIAGONAL"
    cboPatten.AddItem "CROSS"
    cboPatten.AddItem "DIAGCROSS"
    cboPatten.ListIndex = 0
End Sub

Private Function GetDLGName(Optional dShowOpen As Boolean = True, Optional dlgFilter As String = dFilter1, _
Optional dTitle As String = "Open") As String
On Error GoTo OpenErr:

    With CD1
        .CancelError = True
        .DialogTitle = dTitle
        .Filter = dlgFilter
        .InitDir = dInitDir
        .Filename = ""
        
        If (dShowOpen) Then
            .ShowOpen
        Else
            .ShowSave
        End If
        
        'Set InitDir
        dInitDir = GetFilePath(.Filename)
        'Return filename
        GetDLGName = .Filename
    End With
    
    Exit Function
OpenErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Private Sub SetGrips()
    pGrip(0).Left = SrcDc.ScaleWidth
    pGrip(0).Top = (SrcDc.ScaleHeight - 6) \ 2
    pGrip(1).Top = SrcDc.ScaleHeight
    pGrip(1).Left = SrcDc.ScaleWidth
    pGrip(2).Top = SrcDc.ScaleHeight
    pGrip(2).Left = (SrcDc.ScaleWidth - 6) \ 2
End Sub

Private Sub SetScrollbars()
    vBar.Max = (SrcDc.ScaleHeight - pHolder.ScaleHeight) + 18
    hBar.Max = (SrcDc.Width - pHolder.ScaleWidth) + 18
    'Enable/Disable scrollbars.
    hBar.Enabled = (hBar.Max > 0)
    vBar.Enabled = (vBar.Max > 0)
End Sub

Private Sub cboBrush_Click()
    SrcBrush1.Picture = LoadPicture(DataDir & "brushs\" & cboBrush.Text)
End Sub

Private Sub cboCloneSize_Click()
    'Set Clone size
    shpClone.Width = Val(cboCloneSize.Text)
    shpClone.Height = Val(cboCloneSize.Text)
End Sub

Private Sub cboDW_Click()
    SrcDc.DrawWidth = Val(cboDW.Text)
End Sub

Private Sub cboFillType_Click()
    lblPatten.Visible = (Not cboFillType.ListIndex = 0)
    cboPatten.Visible = (Not cboFillType.ListIndex = 0)
    
    If (cboFillType.ListIndex = 1) Then
        lblPatten.Caption = "Patten:"
        'Load pattens
        Call LoadPattens
    End If
    
    If (cboFillType.ListIndex = 2) Then
        lblPatten.Caption = "Hatch:"
        'Load Hatchs
        Call LoadHatchs
    End If
    
End Sub

Private Sub cboPatten_Click()
Dim xFile As String
    'Load patten
    If (cboFillType.ListIndex = 1) Then
        'Patten Filename
        xFile = DataDir & "pattens\" & cboPatten.Text
        'Load the patten
        ImgPatten.Picture = LoadPicture(xFile)
        xFile = vbNullString
    End If
End Sub

Private Sub cboStyle_Click()
    SrcDc.DrawStyle = cboStyle.ListIndex
End Sub

Private Sub cmdTool_Click(Index As Integer)
On Error Resume Next

    'Set mouse cursor
    Call SetCursor(cmdtool(Index).Tag)
    'Hide/show tool options
    pFillType.Visible = (Index = 12)
    pBrush.Visible = (Index = 16)
    pErase.Visible = (Index = 17)
    pClone.Visible = (Index = 19)
    pGrad.Visible = (Index = 20)
    pTextStyle.Visible = (Index = 21)
    lblArrowSize.Visible = (Index = 23)
    cboArrowSize.Visible = (Index = 23)
    
    pDrawWidth.Visible = (Not Index = 0) And (Not Index = 12) _
    And (Not Index = 16) And (Not Index = 18) And (Not Index = 19) _
    And (Not Index = 20) And (Not Index = 21)
    
    lblStyle.Visible = (Not Index = 1) And (Not Index = 14)
    cboStyle.Visible = (Not Index = 1) And (Not Index = 14)
    
    DrawTool = Index
    'Hide selection tool if visuable
    shpSelect.Visible = False
    'Reset clone state
    IsCloneing = False
    shpClone.Visible = False
    mX = 0
    
    If (DrawTool = tClone) Then MsgBox "Use Alt+Mouse1 to define a new area to clone.", vbInformation, frmmain.Caption
    If (DrawTool = tText) Then MsgBox "Click on the area were you like to place the text.", vbInformation, frmmain.Caption
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Swap the colors
    If (KeyAscii = 120) Then
        Call dColorPallete1.SwapColors
    End If
End Sub

Private Sub Form_Load()
Dim Cnt As Integer

    dInitDir = FixPath(App.Path)
    DataDir = FixPath(App.Path) & "Data\"
    mAlphaLevel = 128
    
    ButtonPress = vbCancel
    
    pGrip(1).Picture = pGrip(0).Picture
    pGrip(2).Picture = pGrip(0).Picture
    
    Call SetGrips
    'Update statusbar text
    sBar1.Panels(3).Text = SrcDc.Width & "x" & SrcDc.Height
    'Set Fill Option
    cboFillType.AddItem "Color"
    cboFillType.AddItem "Patten"
    cboFillType.AddItem "Hatch"
    'Set drawing styles
    cboStyle.AddItem "Soild"
    cboStyle.AddItem "Dash"
    cboStyle.AddItem "Dot"
    cboStyle.AddItem "Dash-Dot"
    cboStyle.AddItem "Dash-Dot-Dot"
    'Set erase types
    cboEraseType.AddItem "Square"
    cboEraseType.AddItem "Circle"
    'Add Gradient
    cboGrade.AddItem "Left To Right"
    cboGrade.AddItem "Top To Bottom"
    'Add Text Style
    cboTextStyle.AddItem "Normal"
    cboTextStyle.AddItem "3D Effect"
    cboTextStyle.AddItem "OutLine"
    
    'Load Draw Widths
    For Cnt = 1 To 100
        cboDW.AddItem Cnt
        cboWidth.AddItem Cnt
        'Add arrow sizes
        cboArrowSize.AddItem Cnt
    Next Cnt
    
    'Add some clone sizes
    For Cnt = 1 To 30
        cboCloneSize.AddItem Cnt
    Next Cnt
    
    'Set Index's
    cboFillType.ListIndex = 0
    cboEraseType.ListIndex = 0
    cboDW.ListIndex = 0
    cboStyle.ListIndex = 0
    cboGrade.ListIndex = 0
    cboTextStyle.ListIndex = 0
    cboWidth.ListIndex = 9
    cboCloneSize.ListIndex = 5
    cboArrowSize.ListIndex = 9
    'Load brushs
    Call LoadBrushs
    Call cmdTool_Click(1)
    Call MnuDQue_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    'Resize and position controls
    pHolder.Width = (frmmain.ScaleWidth - pHolder.Left)
    pHolder.Height = (frmmain.ScaleHeight - sBar1.Height - pHolder.Top)
    pTools.Height = pHolder.ScaleHeight + 2
    
    'Position Scrollbars
    vBar.Left = pHolder.ScaleWidth - vBar.Width
    vBar.Height = (pHolder.ScaleHeight) - 18
    hBar.Top = (pHolder.ScaleHeight - hBar.Height)
    hBar.Width = (pHolder.ScaleWidth) - 18
    
    'Position Spacer
    pSpacer.Left = (pHolder.ScaleWidth - pSpacer.Width)
    pSpacer.Top = (pHolder.ScaleHeight - pSpacer.Height)
    
    'Setup scrollbars
    Call SetScrollbars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmCustFltr
    'Close down GDI
    Call GDIShutdown
    Set frmCustFltr = Nothing
    Set frmNew = Nothing
    Set frmScale = Nothing
    Set frmStroke = Nothing
    Set frmText = Nothing
    Set frmRgb = Nothing
    Set frmBright = Nothing
    Set frmMask = Nothing
    Set frmGrade = Nothing
    Set frmabout = Nothing
    Set frmmain = Nothing
End Sub

Private Sub hBar_Change()
    SrcDc.Left = -hBar.Value
End Sub

Private Sub hBar_Scroll()
    Call hBar_Change
End Sub

Private Sub mnuAbout_Click()
   frmabout.Show vbModal, frmmain
End Sub

Private Sub mnuAddNoise_Click()
Dim iVal As Integer
    iVal = Val(InputBox("Enter the noise value.", "Noise", 16))
    'Invert the image's colors
    Call Noise(SrcDc, iVal)
End Sub

Private Sub mnuAdjust_Click()
    frmRgb.Show vbModal, frmmain
    
    If (ButtonPress = vbOK) Then
        'Adjust image colors
        Call AdjustRGB(SrcDc, TmpRgb.R, TmpRgb.G, TmpRgb.B)
    End If
    
    ButtonPress = vbCancel
End Sub

Private Sub mnuAll_Click()
    Call InvertImage(SrcDc, 0)
End Sub

Private Sub mnuBlackWhite_Click()
    'Turn image black and white
    Call BlackAndWhite(SrcDc)
End Sub

Private Sub mnuBlue_Click()
    'Remove Blue
    Call RemoveRGB(SrcDc, 2)
End Sub

Private Sub mnuBlue1_Click()
    Call InvertImage(SrcDc, 3)
End Sub

Private Sub mnuBlur_Click()
    'Blue image
    Call BlurImage(SrcDc)
End Sub

Private Sub mnuBlureMore_Click()
    Call mnuBlur_Click
    Call mnuBlur_Click
    Call mnuBlur_Click
End Sub

Private Sub mnuBoth_Click()
    Call FlipImage(SrcDc, 2)
End Sub

Private Sub mnuBright_Click()
    frmBright.Show vbModal, frmmain
    If (ButtonPress = vbYes) Then
        'Change image brightness
        Call LightenDarken(SrcDc, mBrightness)
    End If
    
    ButtonPress = vbCancel
End Sub

Private Sub mnuCBurn_Click()
    Call ColorBurn(SrcDc)
End Sub

Private Sub mnuClear_Click()
    'Clear the image
    Set SrcDc.Picture = Nothing
End Sub

Private Sub mnuContrast_Click()
Dim iVal As Integer
    'Contrast
    iVal = Val(InputBox("Enter contrast value:", "Contrast", 1.5))
    If (iVal) > 0 Then
        Call Contrast(SrcDc, iVal)
    End If
End Sub

Private Sub mnuCopy_Click()
Dim cOffset As Integer
    cOffset = 1
    If (shpSelect.Visible) Then
        CanPaste = False
        If (CopySelect = True) Then cOffset = 0
        'Copy image selection
        Call CutCopyImage(True, cOffset, False)
        'Hide selection
        shpSelect.Visible = False
    End If
End Sub

Private Sub mnuCrop_Click()
    If (shpSelect.Visible) Then
        'Crop selected area of image
        With pSrc2
            .Width = shpSelect.Width
            .Height = shpSelect.Height
            'Place src picture ontp psrc2
            BitBlt .hdc, 0, 0, pSrc2.Width, pSrc2.Height, SrcDc.hdc, shpSelect.Left, shpSelect.Top, vbSrcCopy
            .Refresh
            'Resize srcdc
            SrcDc.Width = .Width
            SrcDc.Height = .Height
            'Remove Srcdc picture data
            Set SrcDc.Picture = Nothing
            'Place pSrc2 data onto Srcdc
            BitBlt SrcDc.hdc, 0, 0, .Width, .Height, .hdc, 0, 0, vbSrcCopy
            'Refresh srcdc
            SrcDc.Refresh
            Set .Picture = Nothing
            'Hide selection
            shpSelect.Visible = False
            'Set grips
            Call SetGrips
            'Setup scrollbars
            Call SetScrollbars
        End With
    End If
End Sub

Private Sub mnuCustom_Click()
    frmCustFltr.Show vbModal, frmmain
    If (ButtonPress = vbOK) Then
        Call CustomFlilter(SrcDc)
    End If
    
    ButtonPress = vbCancel
    
End Sub

Private Sub mnuCut_Click()
Dim cOffset As Integer
    
    cOffset = 1
    
    If (shpSelect.Visible) Then
        CanPaste = False
        
        If (CopySelect = True) Then cOffset = 0
        
        'Copy image selection
        Call CutCopyImage(True, cOffset)
        'Hide selection
        shpSelect.Visible = False
    End If
End Sub

Private Sub mnuDiffuse_Click()
    'Diffuse
    Call Diffuse(SrcDc)
End Sub

Private Sub MnuDQue_Click()
    MnuDQue.Checked = (Not MnuDQue.Checked)
End Sub

Private Sub mnuExit_Click()
    Unload frmmain
End Sub

Private Sub mnuFlipHoz_Click()
    Call FlipImage(SrcDc, 1)
End Sub

Private Sub mnuFlipV_Click()
    Call FlipImage(SrcDc, 0)
End Sub

Private Sub mnuGamma_Click()
Dim iVal As Integer
    'Generate a grid
    iVal = Val(InputBox("Enter gamma value:", "Gamma", 2.5))
    If (iVal) > 0 Then
        Call Gamma(SrcDc, iVal)
    End If
End Sub

Private Sub mnuGrade_Click()
    'Show Gradient form
    frmGrade.Show vbModal, frmmain
    
    If (ButtonPress = vbOK) Then
        Call GradientAlpha(SrcDc, pSrc2, mAlphaLevel, dColorPallete1.Color1, dColorPallete1.Color2, AlphaDir)
    End If
    
    ButtonPress = vbCancel
End Sub

Private Sub mnuGray_Click()
    'Grayscale image
    Call GrayImage(SrcDc)
End Sub

Private Sub mnuGreen_Click()
    'Remove Green
    Call RemoveRGB(SrcDc, 1)
End Sub

Private Sub mnuGreen1_Click()
    Call InvertImage(SrcDc, 2)
End Sub

Private Sub mnuGrid_Click()
Dim iVal As Integer
    'Generate a grid
    iVal = Val(InputBox("Enter a grid size:", "Grid Generator", 10))
    If (iVal) > 0 Then
        Call DrawGrid(SrcDc, iVal)
    End If
End Sub

Private Sub mnuHighlight_Click()
    frmHighlight.Show vbModal, frmmain
    If (ButtonPress = vbOK) Then
        'Do highlight
        Call Highlight(SrcDc, TmpRgb.R, TmpRgb.G, TmpRgb.B)
    End If
  
    ButtonPress = vbCancel
    
End Sub

Private Sub mnuImboss_Click()
    'Imboss effect
    Call Flatten(SrcDc, -128)
End Sub

Private Sub mnuMask_Click()
    frmMask.Show vbModal, frmmain
    
    If (ButtonPress = vbOK) Then
        'Appply the image mask
        Call ApplyImageMask(mMaskFile)
    End If
    
    ButtonPress = vbCancel
End Sub

Private Sub mnuMelt_Click()
    'Melt effect
    Call MeltEffect(SrcDc)
End Sub

Private Sub mnuMirror_Click()
    Call Mirror(SrcDc, 0)
End Sub

Private Sub mnuMRight_Click()
    'Mirror right
    Call Mirror(SrcDc, 1)
End Sub

Private Sub mnuNew_Click()
    frmNew.Show vbModal, frmmain
    'Check if ok button was pressed
    If (ButtonPress = vbOK) Then
        'setup the new image
        Set SrcDc.Picture = Nothing
        SrcDc.Width = mImage.iWidth
        SrcDc.Height = mImage.iHeight
        SrcDc.BackColor = mImage.iBackColor
        MnuDQue.Checked = False
        'Hide select tool is visable
        If (shpSelect.Visible) Then
            shpSelect.Visible = False
        End If
        'Position grips
        Call SetGrips
        'Resize
        Call Form_Resize
    End If
    
    'Update statusbar text
    sBar1.Panels(3).Text = SrcDc.Width & "x" & SrcDc.Height
    ButtonPress = vbCancel
    
End Sub
Private Sub mnuOpen_Click()
Dim lFile As String
    
    'Hide selection if visable
    If (shpSelect.Visible) Then
        shpSelect.Visible = False
    End If
    
    lFile = GetDLGName()
    
    If Len(lFile) Then
        'Init GDI
        Call InitGDIPlus
        Call GDILoadPicture(lFile, SrcDc)
        'SrcDc.Picture = LoadPicture(lFile)
        SrcDc.BackColor = vbWhite
        MnuDQue.Checked = False
        'Position grips
        Call SetGrips
        'Resize
        Call Form_Resize
        'Shut down GDI
        Call GDIShutdown
    End If
    
    DrawTool = 1
    
    'Update statusbar text
    sBar1.Panels(3).Text = SrcDc.Width & "x" & SrcDc.Height
End Sub

Private Sub mnuPaste_Click()
    'Get image data from clipboard
    pSrc2.Picture = Clipboard.GetData(vbCFBitmap)
    CanPaste = True
    DrawTool = -1
End Sub

Private Sub mnuPixel_Click()
Dim iVal As Integer
    'Generate a grid
    iVal = Val(InputBox("Enter pixel size:", "Pixel Effect", 5.2))
    If (iVal) > 0 Then
        Call PixelsEffect(SrcDc, iVal)
    End If
End Sub

Private Sub mnuPlaster_Click()
    'Plaster effect
    Call Flatten(SrcDc, -210)
End Sub

Private Sub mnuRed_Click()
    'Remove Red
    Call RemoveRGB(SrcDc, 0)
End Sub

Private Sub mnuRed1_Click()
    Call InvertImage(SrcDc, 1)
End Sub

Private Sub mnuReplaceC_Click()
    'Do replace color
    Call ReplaceColor(SrcDc, dColorPallete1.Color1, dColorPallete1.Color2)
End Sub

Private Sub mnuSave_Click()
Dim lFile As String
Dim MineT As String
Dim iRet As Boolean
    lFile = GetDLGName(False, dFilter2, "Save")
    
    If Len(lFile) Then
        'Save picture
        'This code is added to fix a bug while saveing. Pictures are now saved in the correct size.
        'Resize pSrc2 to hold the image
        pSrc2.Width = SrcDc.Width
        pSrc2.Height = SrcDc.Height
        'Copy image to pSrc2
        Call BitBlt(pSrc2.hdc, 0, 0, pSrc2.Width, pSrc2.Height, SrcDc.hdc, 0, 0, vbSrcCopy)
        'Get Mime type
        Select Case LCase(GetFileExt(lFile))
            Case "png": MineT = "image/png"
            Case "jpg": MineT = "image/jpeg"
            Case "gif": MineT = "image/gif"
            Case "bmp": MineT = "BMP"
        End Select
        'Check mine type
        If (MineT = "BMP") Then
            'Save normal bitmap
            Call SavePicture(pSrc2.image, lFile)
        Else
            'Init GDI
            Call InitGDIPlus
            'Save a copy of the image cos it need to be a bitmap sure there a better way of doing this
            'but for now it seems to work
            Call SavePicture(pSrc2.image, lFile)
            'reload the image
            SrcDc.Picture = LoadPicture(lFile)
            'Delete the file
            Call Kill(lFile)
            'Now we can use the GDI Plus to save the picture
            iRet = GDISavePicture(SrcDc, lFile, MineT)
            'shutdown gdi
            Call GDIShutdown
        End If
        
        'Call SavePicture(pSrc2.image, lFile)
        Set pSrc2.Picture = Nothing
    End If
    
End Sub

Private Sub mnuScale_Click()
    mImage.iHeight = SrcDc.Height
    mImage.iWidth = SrcDc.Width
    frmScale.Show vbModal, frmmain
    'Check buttonpressed
    If (ButtonPress = vbOK) Then
        'Scale the image
        Call ScaleImage(SrcDc, pSrc2, mImage.iWidth, mImage.iHeight)
        'Set grips
        Call SetGrips
        'Setup scrollbars
        Call SetScrollbars
    End If
    
    ButtonPress = vbCancel
    
End Sub

Private Sub mnuSelAll_Click()
    'This selects the whole image
    CopySelect = True
    shpSelect.Move 0, 0, SrcDc.ScaleWidth, SrcDc.ScaleHeight
    shpSelect.Visible = True
    SrcDc.SetFocus
End Sub

Private Sub mnuSharpen_Click()
On Error GoTo ValErr:
    mSharpen = CSng(InputBox("Enter sharpen value.", "Sharpen", 1.2))
    'Sharpen Image
    Call Sharpen(SrcDc, mSharpen)
    Exit Sub
ValErr:
    MsgBox Err.Description, vbInformation, "Sharpen"
End Sub

Private Sub mnuSharpMore_Click()
    'Sharpen agian
    Call Sharpen(SrcDc, mSharpen)
End Sub

Private Sub mnuSnow_Click()
Dim iVal As Integer
    'Snow effect
    iVal = Val(InputBox("Enter a snow size:", "Grid Generator", 32))
    If (iVal) > 0 Then
        Call Snow(SrcDc, iVal)
    End If
End Sub

Private Sub mnuStroke_Click()
    frmStroke.Show vbModal, frmmain
    'Check the button pressed
    If (ButtonPress = vbOK) Then
        'Stroke the image
        Call StrokeImage(SrcDc, StrokeW, StrokeC)
    End If
    
    ButtonPress = vbCancel
End Sub

Private Sub mnuSwapRGB_Click()
    frmRgbSwap.Show vbModal, frmmain
    If (ButtonPress = vbOK) Then
        'Swap RGB Channels
        Call SwapRGB(SrcDc, mRGBSwapIndex)
    End If
    
    ButtonPress = vbCancel
End Sub

Private Sub mnuXorC_Click()
Dim iVal As Integer
    'XOR Color
    iVal = Val(InputBox("Enter a value:", "XORColor", 10))
    If (iVal) > 0 Then
        Call XorColor(SrcDc, iVal)
    End If
End Sub

Private Sub pGrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) Then
        MoveX = X
        MoveY = Y
        IsMoveing = True
    End If
End Sub

Private Sub pGrip_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (IsMoveing) Then
        
        'This allows the user to resize the drawing area by draging the grips
        Select Case Index
            Case 0
                pGrip(0).Left = (pGrip(0).Left - (MoveX - X) \ Screen.TwipsPerPixelX)
                If pGrip(0).Left < 1 Then pGrip(0).Left = 1
                SrcDc.Width = pGrip(0).Left
            Case 1
                pGrip(1).Top = (pGrip(1).Top - (MoveY - Y) \ Screen.TwipsPerPixelY)
                pGrip(1).Left = (pGrip(1).Left - (MoveX - X) \ Screen.TwipsPerPixelX)
                If pGrip(1).Left < 1 Then pGrip(1).Left = 1
                If pGrip(1).Top < 1 Then pGrip(1).Top = 1
                SrcDc.Width = pGrip(1).Left
                SrcDc.Height = pGrip(1).Top
            Case 2
                pGrip(2).Top = (pGrip(2).Top - (MoveY - Y) \ Screen.TwipsPerPixelY)
                If pGrip(2).Top < 1 Then pGrip(2).Top = 1
                SrcDc.Height = pGrip(2).Top
        End Select
        
        'Position the grips
        Call SetGrips
        'Setup scrollbars
        Call SetScrollbars
        'Update statusbar text
        sBar1.Panels(3).Text = SrcDc.Width & "x" & SrcDc.Height
    End If
    
End Sub
Private Sub pTop_Resize()
    lnTop1.Width = pTop.ScaleWidth
    lnTop2.Width = pTop.ScaleWidth
End Sub

Private Sub SrcDc_DblClick()
    'Cancel polygone drawing tool
    IsDown = False
End Sub

Private Sub SrcDc_KeyUp(KeyCode As Integer, Shift As Integer)
    If (shpSelect.Visible) Or (CopySelect = True) And (KeyCode = vbKeyDelete) Then
        Call mnuCut_Click
    End If
End Sub

Private Sub SrcDc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hBrush As Long
Dim iRet As Long

    With mDrawPos
    
        'Sets the drawing forecolor
        If (Button = 1) Then SrcDc.ForeColor = dColorPallete1.Color1
        If (Button = 2) Then SrcDc.ForeColor = dColorPallete1.Color2
        
        .StartX = X
        .StartY = Y
            
        .OldX = .StartX
        .OldX = .StartY
        CanDraw = True
        
        'Triangle tool
        If (DrawTool = tTriangle) Then
            Select Case mX
                Case 0
                    SrcDc.PSet (X, Y)
                    .X1 = X
                    .Y1 = Y
                    mX = 1
                Case 1
                    SrcDc.Line -(X, Y)
                    mX = 2
                Case 2
                    SrcDc.Line -(X, Y)
                    SrcDc.Line -(.X1, .Y1)
                    mX = 0
            End Select
        End If
        
        'Pencil tool
        If (DrawTool = tPencil) Then
            Call SrcDc_MouseMove(Button, Shift, X, Y)
        End If
        
        'Fill can tool
        If (DrawTool = tFillCan) Then
            SrcDc.FillStyle = vbFSSolid
            SrcDc.FillColor = SrcDc.ForeColor
            
            'Check if filling with a patten
            If (cboFillType.ListIndex = 1) Then
                'Create the brush
                hBrush = CreatePatternBrush(ImgPatten.Picture.Handle)
                'Place brush onto the object
                iRet = SelectObject(SrcDc.hdc, hBrush)
            End If
            'Check if filling with hatch brush
            If (cboFillType.ListIndex = 2) Then
                hBrush = CreateHatchBrush(cboPatten.ListIndex, SrcDc.ForeColor)
                'Place brush onto the object
                iRet = SelectObject(SrcDc.hdc, hBrush)
            End If
            
            iRet = ExtFloodFill(SrcDc.hdc, .StartX, .StartY, SrcDc.Point(.StartX, .StartY), FLOODFILLSURFACE)
            
            'Destory Brush
            If (cboFillType.ListIndex = 1) Then
                iRet = DeleteObject(hBrush)
            End If
        End If
        
        'Spray Tool
        If (DrawTool = tSpray) Then
            Call SrcDc_MouseMove(Button, Shift, X, Y)
        End If
        
        'Erase tool
        If (DrawTool = tErase) Then
            Call SrcDc_MouseMove(Button, Shift, X, Y)
        End If
        
        'Brush tool
        If (DrawTool = tBrush) Then
            Call SrcDc_MouseMove(Button, Shift, X, Y)
        End If
        
        'Poly Tool
        If (DrawTool = tCurve) Then
            .OldX = X
            .OldY = Y
        End If
        
        'Color picker
        If (DrawTool = tColorPicker) Then
            If (Button = vbLeftButton) Then dColorPallete1.Color1 = SrcDc.Point(X, Y)
            If (Button = vbRightButton) Then dColorPallete1.Color2 = SrcDc.Point(X, Y)
        End If

        'Polygone tool
        If (DrawTool = tPoly) Then
            If (Not IsDown) Then
                .X1 = X
                .Y1 = Y
                SrcDc.PSet (.X1, .Y1), SrcDc.ForeColor
                IsDown = True
            Else
                SrcDc.Line (.X1, .Y1)-(X, Y), SrcDc.ForeColor
            End If
        End If
        
        'Clone tool
        If (DrawTool = tClone) And (Shift = 4) Then
            Call SetCursor(0)
            IsCloneing = True
            mDrawPos.X1 = X
            mDrawPos.Y1 = Y
            'Center clone tool
            shpClone.Left = (X - shpClone.Width \ 2)
            shpClone.Top = (Y - shpClone.Height \ 2)
            shpClone.Visible = True
        End If
        
        'Select tool
        If (DrawTool = tSelect) Or (CopySelect = True) Then
            shpSelect.Visible = False
            CopySelect = False
        End If

        SrcDc.AutoRedraw = False
        SrcDc.FillStyle = vbFSTransparent
    End With
End Sub

Private Sub SrcDc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cnt As Integer
Dim X1 As Integer
Dim Y1 As Integer
Dim iRet As Long
Dim rc As RECT

    If (CanDraw = True) Then
        
        'Paste picture
        If (CanPaste) Then
            SrcDc.Refresh
            If (MnuDQue.Checked) Then
                TransparentBlt SrcDc.hdc, (X - pSrc2.Width \ 2), (Y - pSrc2.Height \ 2), pSrc2.Width, pSrc2.Height, pSrc2.hdc, 0, 0, pSrc2.Width, pSrc2.Height, SrcDc.BackColor
            Else
                BitBlt SrcDc.hdc, (X - pSrc2.Width \ 2), (Y - pSrc2.Height \ 2), pSrc2.Width, pSrc2.Height, pSrc2.hdc, 0, 0, vbSrcCopy
            End If
        End If
        
        Select Case DrawTool
            Case tPencil
                'Freehand drawing tool
                SrcDc.AutoRedraw = True
                SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y)
                mDrawPos.StartX = X
                mDrawPos.StartY = Y
            Case tline
                'Line tool
                SrcDc.Refresh
                SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y)
            Case tStar
                'Star tool
                SrcDc.AutoRedraw = True
                SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y)
            Case tBox
                'Box tool
                SrcDc.Refresh
                SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y), , B
            Case tBoxF
                'Filled box tool
                SrcDc.Refresh
                SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y), , BF
            Case tCircle
                'Circle tool
                SrcDc.Refresh
                SrcDc.Circle (mDrawPos.StartX, mDrawPos.StartY), Sqr((X - mDrawPos.StartX) ^ 2 + (Y - mDrawPos.StartY) ^ 2)
            Case tCircleF
                'Circle tool filled
                SrcDc.FillColor = SrcDc.ForeColor
                SrcDc.FillStyle = 0
                SrcDc.Refresh
                SrcDc.Circle (mDrawPos.StartX, mDrawPos.StartY), Sqr((X - mDrawPos.StartX) ^ 2 + (Y - mDrawPos.StartY) ^ 2)
            Case tEllipse
                'Ellipse tool
                SrcDc.Refresh
                iRet = Ellipse(SrcDc.hdc, mDrawPos.StartX, mDrawPos.StartY, X, Y)
            Case tEllipseF
                'Ellipse tool filled
                SrcDc.Refresh
                SrcDc.FillColor = SrcDc.ForeColor
                SrcDc.FillStyle = 0
                iRet = Ellipse(SrcDc.hdc, mDrawPos.StartX, mDrawPos.StartY, X, Y)
            Case tRoundRect
                'Round rectangle tool
                SrcDc.Refresh
                iRet = RoundRect(SrcDc.hdc, mDrawPos.StartX, mDrawPos.StartY, X, Y, 12, 12)
            Case tRoundRectF
                'Round rectangle tool filled
                SrcDc.Refresh
                SrcDc.FillColor = SrcDc.ForeColor
                SrcDc.FillStyle = 0
                iRet = RoundRect(SrcDc.hdc, mDrawPos.StartX, mDrawPos.StartY, X, Y, 12, 12)
            Case tSpray
                'Spray tool
                SrcDc.AutoRedraw = True
                For Cnt = 0 To 10
                    X1 = Int(Rnd * SrcDc.DrawWidth * 10) - 5
                    Y1 = Int(Rnd * SrcDc.DrawWidth * 10) - 5
                    SrcDc.PSet (mDrawPos.StartX + X1, mDrawPos.StartY + Y1)
                    mDrawPos.StartX = X
                    mDrawPos.StartY = Y
                Next Cnt
            Case tCurve
                'Poly Tool
                SrcDc.AutoRedraw = True
                SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y)
                mDrawPos.StartX = X
                mDrawPos.StartY = Y
            Case tErase
                'Erase tool
                mDrawPos.X1 = (X - Val(cboWidth.Text) \ 2)
                mDrawPos.Y1 = (Y - Val(cboWidth.Text) \ 2)
                If (cboEraseType.ListIndex = 0) Then
                    'Draw with a square
                    SrcDc.Line (mDrawPos.X1, mDrawPos.Y1)-(mDrawPos.X1 + Val(cboWidth.Text), mDrawPos.Y1 + Val(cboWidth.Text)), SrcDc.BackColor, BF
                Else
                    'Draw with a circle
                    SrcDc.FillStyle = vbFSSolid
                    SrcDc.FillColor = SrcDc.BackColor
                    SrcDc.Circle (mDrawPos.X1 + Val(cboWidth.Text) \ 2, mDrawPos.Y1 + Val(cboWidth.Text) \ 2), Val(cboWidth.Text), SrcDc.BackColor
                End If
                SrcDc.AutoRedraw = True
            Case tColorPicker
                'Color picker
                If (Button = vbLeftButton) Then dColorPallete1.Color1 = SrcDc.Point(X, Y)
                If (Button = vbRightButton) Then dColorPallete1.Color2 = SrcDc.Point(X, Y)
            Case tGradient1
                SrcDc.Refresh
                'Gradient Tool
                iRet = SetRect(rc, mDrawPos.StartX, mDrawPos.StartY, X, Y)
                Call GDI_GradientFill(SrcDc.hdc, rc, dColorPallete1.Color1, dColorPallete1.Color2, cboGrade.ListIndex)
            Case tPoly
                'Polygone tool
                SrcDc.Refresh
                SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y)
            Case tBrush
                'Brush tool
                SrcDc.AutoRedraw = True
                Call DrawBrush
                iRet = TransparentBlt(SrcDc.hdc, (X - SrcBrush2.Width \ 2), (Y - SrcBrush2.Height \ 2), SrcBrush2.Width, SrcBrush2.Height, _
                SrcBrush2.hdc, 0, 0, SrcBrush2.Width, SrcBrush2.Height, RGB(255, 0, 255))
                SrcDc.Refresh
            Case tClone
                'Clone tool
                If (IsCloneing) Then
                    TmpX = (X - mDrawPos.X1)
                    TmpY = (Y - mDrawPos.Y1)
                    IsCloneing = False
                Else
                    SrcDc.AutoRedraw = True
                    'Do the cloneing
                    Call BitBlt(SrcDc.hdc, X, Y, shpClone.Width, shpClone.Height, SrcDc.hdc, X - TmpX, Y - TmpY, vbSrcCopy)
                    'Position clone tool
                    shpClone.Top = (Y - TmpY)
                    shpClone.Left = (X - TmpX)
                    SrcDc.Refresh
                End If
            Case tSelect
                'Selection tool
                
                If (Y < 0) Then Y = 0
                If (X < 0) Then X = 0
                If (X > SrcDc.ScaleWidth) Then X = (SrcDc.ScaleWidth - 1)
                If (Y > SrcDc.ScaleHeight) Then Y = (SrcDc.ScaleHeight - 1)
                
                If (mDrawPos.StartX > X) Then
                    shpSelect.Left = X
                    shpSelect.Width = (mDrawPos.StartX - X)
                End If
                
                If (mDrawPos.StartX < X) Then
                    shpSelect.Left = mDrawPos.StartX
                    shpSelect.Width = (X - mDrawPos.StartX)
                End If
                
                If (mDrawPos.StartY > Y) Then
                    shpSelect.Top = Y
                    shpSelect.Height = (mDrawPos.StartY - Y)
                End If
                
                If (mDrawPos.StartY < Y) Then
                    shpSelect.Top = mDrawPos.StartY
                    shpSelect.Height = (Y - mDrawPos.StartY)
                End If
                
                'Display selection
                shpSelect.Visible = True
            Case tArrow
                'Arrow tool
                SrcDc.Refresh
                Call DrawArrow(SrcDc, mDrawPos.StartX, mDrawPos.StartY, X, Y, Val(cboArrowSize.Text))
        End Select
    End If
    
    'Update statusbar text
    sBar1.Panels(2).Text = X & "," & Y
End Sub

Private Sub SrcDc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rc As RECT
Dim iRet As Long

    SrcDc.FillStyle = vbFSTransparent
    SrcDc.AutoRedraw = True
    
    'Check if can paste
    If (CanPaste) Then
        'Paste the image
        If (MnuDQue.Checked) Then
            TransparentBlt SrcDc.hdc, (X - pSrc2.Width \ 2), (Y - pSrc2.Height \ 2), pSrc2.Width, pSrc2.Height, _
            pSrc2.hdc, 0, 0, pSrc2.Width, pSrc2.Height, SrcDc.BackColor
        Else
            BitBlt SrcDc.hdc, (X - pSrc2.Width \ 2), (Y - pSrc2.Height \ 2), pSrc2.Width, pSrc2.Height, pSrc2.hdc, 0, 0, vbSrcCopy
        End If
        'Cancel Paste
        CanPaste = False
    End If
    
    Select Case DrawTool
        Case tline
            'Line tool
            SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y)
        Case tBox
            'Box tool
            SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y), , B
        Case tBoxF
            'Filled box tool
            SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y), , BF
        Case tCircle
            'Circle tool
            SrcDc.Circle (mDrawPos.StartX, mDrawPos.StartY), Sqr((X - mDrawPos.StartX) ^ 2 + (Y - mDrawPos.StartY) ^ 2)
        Case tCircleF
            'Filled circle tool
            SrcDc.FillStyle = 0
            SrcDc.Circle (mDrawPos.StartX, mDrawPos.StartY), Sqr((X - mDrawPos.StartX) ^ 2 + (Y - mDrawPos.StartY) ^ 2)
        Case tEllipse
            'Ellipse tool
            iRet = Ellipse(SrcDc.hdc, mDrawPos.StartX, mDrawPos.StartY, X, Y)
        Case tEllipseF
            'Filled ellipse tool
            SrcDc.FillStyle = 0
            iRet = Ellipse(SrcDc.hdc, mDrawPos.StartX, mDrawPos.StartY, X, Y)
        Case tRoundRect
            'RoundRect tool
            iRet = RoundRect(SrcDc.hdc, mDrawPos.StartX, mDrawPos.StartY, X, Y, 12, 12)
        Case tRoundRectF
            'Filled round rect tool
            SrcDc.FillStyle = 0
            iRet = RoundRect(SrcDc.hdc, mDrawPos.StartX, mDrawPos.StartY, X, Y, 12, 12)
        Case tCurve
            'Curve tool
            SrcDc.Line (mDrawPos.OldX, mDrawPos.OldY)-(X, Y)
        Case tGradient1
            'Draw Gradient rectange
            iRet = SetRect(rc, mDrawPos.StartX, mDrawPos.StartY, X, Y)
            Call GDI_GradientFill(SrcDc.hdc, rc, dColorPallete1.Color1, dColorPallete1.Color2, cboGrade.ListIndex)
        Case tPoly
            'Polygone tool
            SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y)
            SrcDc.Line (mDrawPos.StartX, mDrawPos.StartY)-(mDrawPos.X1, mDrawPos.Y1)
            mDrawPos.X1 = X
            mDrawPos.Y1 = Y
        Case tText
             'Text Tool
            frmText.Show vbModal, frmmain
            
            If (ButtonPress = vbOK) Then
                With SrcDc
                    .AutoRedraw = True
                    .FontName = mTextProp.ffName
                    .FontSize = mTextProp.fSize
                    .FontBold = False
                    .FontItalic = False
                    
                    'Set Font Styles
                    Select Case mTextProp.fStyle
                        Case 1
                            .FontBold = True
                        Case 2
                            .FontItalic = True
                        Case 3
                            .FontBold = True
                            .FontItalic = True
                    End Select
                    
                    'Check if using normal text
                    If (cboTextStyle.ListIndex = 0) Then
                        'Update drawing with Text
                        iRet = TextOut(.hdc, X, Y, mTextProp.fText, Len(mTextProp.fText))
                    End If
                    'Check if using 3D Effect
                    If (cboTextStyle.ListIndex = 1) Then
                        iRet = TextOut(.hdc, X, Y, mTextProp.fText, Len(mTextProp.fText))
                        'Set shade color
                        .ForeColor = dColorPallete1.Color2
                        iRet = TextOut(.hdc, X + 2, Y + 2, mTextProp.fText, Len(mTextProp.fText))
                    End If
                    'Check if using outline effect
                    If (cboTextStyle.ListIndex = 2) Then
                        'Outline the text
                        Call OutLineText(SrcDc, X, Y, mTextProp.fText)
                    End If
                    
                End With
            End If
            CanDraw = False
            ButtonPress = vbCancel
        Case tClone
            Call SetCursor(1)
        Case tArrow
            'Arrow tool
            Call DrawArrow(SrcDc, mDrawPos.StartX, mDrawPos.StartY, X, Y, Val(cboArrowSize.Text))
        End Select
    
        'Cancel drawing
        CanPaste = False
        CanDraw = False
        SrcDc.Refresh

End Sub

Private Sub vBar_Change()
    SrcDc.Top = -vBar.Value
End Sub

Private Sub vBar_Scroll()
    Call vBar_Change
End Sub
