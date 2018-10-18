VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H007F7F7F&
   BorderStyle     =   0  'None
   Caption         =   " Maelstrom CD-Key Tester"
   ClientHeight    =   6885
   ClientLeft      =   4125
   ClientTop       =   -225
   ClientWidth     =   11610
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   11610
   Begin VB.Timer tmrCheckUpdate 
      Enabled         =   0   'False
      Interval        =   450
      Left            =   1560
      Top             =   5880
   End
   Begin MSWinsockLib.Winsock sckCheckUpdate 
      Left            =   1560
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrWaitLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   5400
   End
   Begin VB.Timer tmrCheckBNLS 
      Enabled         =   0   'False
      Left            =   1080
      Top             =   5880
   End
   Begin VB.Timer tmrBenchmark 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   4920
   End
   Begin MSWinsockLib.Winsock sckBNLS 
      Left            =   1080
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox pbMinimize 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   10635
      Picture         =   "frmMain.frx":0CCA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   68
      Top             =   -15
      Width           =   480
   End
   Begin MSWinsockLib.Winsock sckBNCS 
      Index           =   0
      Left            =   600
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox pbQuit 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   11115
      Picture         =   "frmMain.frx":0FEA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   -15
      Width           =   480
   End
   Begin VB.Timer tmrCheckFailed 
      Enabled         =   0   'False
      Index           =   0
      Left            =   600
      Top             =   5880
   End
   Begin VB.Timer tmrReconnect 
      Enabled         =   0   'False
      Index           =   0
      Left            =   600
      Top             =   5400
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   3135
      Left            =   120
      TabIndex        =   25
      Top             =   3600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5530
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":15C5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading CD-Key Profile..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   89
      Top             =   720
      Width           =   11295
   End
   Begin VB.Label lblUpdateLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check for update"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7800
      TabIndex        =   88
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblReloadProxies 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reload Proxies"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   87
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label lblReloadCDKeys 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reload CD-Keys"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9360
      TabIndex        =   86
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblConfig 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Config"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9720
      TabIndex        =   85
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label lblControl 
      BackStyle       =   0  'Transparent
      Caption         =   "Testing Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   81
      Left            =   7320
      TabIndex        =   84
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   7320
      X2              =   11400
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   11400
      X2              =   7320
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   7320
      X2              =   11400
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   7320
      X2              =   11400
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   7320
      X2              =   11400
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   80
      Left            =   10320
      TabIndex        =   83
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Sockets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   79
      Left            =   7320
      TabIndex        =   82
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   78
      Left            =   10440
      TabIndex        =   81
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   77
      Left            =   9240
      TabIndex        =   80
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   64
      Left            =   4080
      TabIndex        =   79
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   86
      Left            =   10440
      TabIndex        =   78
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lblControl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   84
      Left            =   10440
      TabIndex        =   77
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblControl 
      BackStyle       =   0  'Transparent
      Caption         =   "Keys Per Second"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   83
      Left            =   7320
      TabIndex        =   76
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label lblControl 
      BackColor       =   &H007F7F7F&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Elapsed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   82
      Left            =   7320
      TabIndex        =   75
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label lblControl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   98
      Left            =   10560
      TabIndex        =   74
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   97
      Left            =   10560
      TabIndex        =   73
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   96
      Left            =   10560
      TabIndex        =   72
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   95
      Left            =   10560
      TabIndex        =   71
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   10440
      TabIndex        =   70
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblControl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   94
      Left            =   10560
      TabIndex        =   69
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   63
      Left            =   3360
      TabIndex        =   18
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   62
      Left            =   2520
      TabIndex        =   5
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   42
      Left            =   8640
      TabIndex        =   67
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   30
      Left            =   5880
      TabIndex        =   66
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jailed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000A4E1&
      Height          =   255
      Index           =   12
      Left            =   5760
      TabIndex        =   65
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   69
      Left            =   8640
      TabIndex        =   64
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   51
      Left            =   8640
      TabIndex        =   63
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   33
      Left            =   8640
      TabIndex        =   62
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   60
      Left            =   8640
      TabIndex        =   61
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   57
      Left            =   5880
      TabIndex        =   60
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   48
      Left            =   5880
      TabIndex        =   59
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   39
      Left            =   5880
      TabIndex        =   58
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   66
      Left            =   5880
      TabIndex        =   57
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Other Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BCBCBE&
      Height          =   495
      Index           =   14
      Left            =   8400
      TabIndex        =   56
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblControl 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C2BF43&
      Height          =   255
      Index           =   16
      Left            =   9720
      TabIndex        =   54
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   52
      Left            =   9465
      TabIndex        =   53
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   70
      Left            =   9480
      TabIndex        =   52
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   32
      Left            =   7680
      TabIndex        =   51
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   61
      Left            =   9465
      TabIndex        =   50
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   43
      Left            =   9465
      TabIndex        =   49
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   31
      Left            =   6840
      TabIndex        =   48
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   50
      Left            =   7665
      TabIndex        =   47
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   59
      Left            =   7665
      TabIndex        =   46
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   41
      Left            =   7665
      TabIndex        =   45
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   28
      Left            =   4080
      TabIndex        =   44
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   38
      Left            =   5025
      TabIndex        =   43
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   56
      Left            =   5025
      TabIndex        =   42
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   65
      Left            =   5025
      TabIndex        =   41
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   47
      Left            =   5040
      TabIndex        =   40
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   40
      Left            =   6825
      TabIndex        =   39
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   58
      Left            =   6825
      TabIndex        =   38
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   67
      Left            =   6825
      TabIndex        =   37
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   29
      Left            =   5025
      TabIndex        =   36
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   68
      Left            =   7680
      TabIndex        =   35
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblControl 
      BackStyle       =   0  'Transparent
      Caption         =   "Banned"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   13
      Left            =   6720
      TabIndex        =   34
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblControl 
      BackStyle       =   0  'Transparent
      Caption         =   "Voided"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   11
      Left            =   5025
      TabIndex        =   33
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblControl 
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7680
      TabIndex        =   32
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   49
      Left            =   6840
      TabIndex        =   31
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblControl 
      BackStyle       =   0  'Transparent
      Caption         =   "Muted"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   10
      Left            =   4200
      TabIndex        =   30
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "In-Use"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E74516&
      Height          =   255
      Index           =   9
      Left            =   3285
      TabIndex        =   24
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "Perfect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   8
      Left            =   2505
      TabIndex        =   23
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   55
      Left            =   4080
      TabIndex        =   17
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   54
      Left            =   3360
      TabIndex        =   16
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   46
      Left            =   4080
      TabIndex        =   15
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   45
      Left            =   3360
      TabIndex        =   14
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   37
      Left            =   4080
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   36
      Left            =   3360
      TabIndex        =   12
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   34
      Left            =   9465
      TabIndex        =   11
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   27
      Left            =   3360
      TabIndex        =   10
      Top             =   1800
      Width           =   660
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   44
      Left            =   2520
      TabIndex        =   9
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   26
      Left            =   2520
      TabIndex        =   8
      Top             =   1800
      Width           =   780
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   35
      Left            =   2520
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   255
      Index           =   53
      Left            =   2520
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblControl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   72
      Left            =   10560
      TabIndex        =   55
      Top             =   4440
      Width           =   855
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   11
      X1              =   120
      X2              =   11400
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   10
      X1              =   120
      X2              =   11400
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   9
      X1              =   120
      X2              =   11400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   13
      X1              =   120
      X2              =   11400
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   12
      X1              =   120
      X2              =   11400
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblStart 
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7680
      TabIndex        =   29
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scanner Statistics"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   71
      Left            =   7320
      TabIndex        =   28
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Maelstrom CD-Key Tester v%ver by Vector"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   75
      Left            =   10200
      TabIndex        =   22
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   74
      Left            =   9240
      TabIndex        =   21
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   76
      Left            =   10440
      TabIndex        =   20
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblControl 
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "CD-Keys Tested"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   73
      Left            =   7320
      TabIndex        =   19
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblControl 
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "WarCraft III: TFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblControl 
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "Warcraft III"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblControl 
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "WarCraft II: BNE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblControl 
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "Diablo II: LoD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblControl 
      BackColor       =   &H00DFBF26&
      BackStyle       =   0  'Transparent
      Caption         =   "Diablo II Classic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nid As NOTIFYICONDATA ' trayicon variable

Sub minimize_to_tray()
    Me.Hide
    nid.cbSize = Len(nid)
    nid.hwnd = Me.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = Me.Icon ' the icon will be your Form1 project icon
    nid.szTip = PROGRAM_NAME & vbNullChar
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(keyCode, Shift)
End Sub

Private Sub Form_Load()
    Dim top As Long, left As Long, tempValue As String

    AddChat vbYellow, "Welcome to Maelstrom CD-Key Tester v" & PROGRAM_VERSION & " by Vector."
    lblControl(1).Caption = PROGRAM_NAME
  
    tmrWaitLoad.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msg As Long
    Dim sFilter As String
    msg = X / Screen.TwipsPerPixelX
  
    Select Case msg
        Case WM_LBUTTONDOWN
        Case WM_LBUTTONUP
            Me.Show ' show form
            Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
            rtbChat.SetFocus
        Case WM_LBUTTONDBLCLK
        Case WM_RBUTTONDOWN
        Case WM_RBUTTONUP
        Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub lblConfig_Click()
    frmConfig.Show
    lblStart.Enabled = False
End Sub

Private Sub lblControl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    moveEntireForm Me, Button
End Sub

Public Sub lblStart_EmulateClick()
    Call lblStart_Click
End Sub

Private Sub lblReloadCDKeys_Click()
    wipeCDKeysFromTesting
    loadCDKeys

    AddChat vbWhite, totalNonExpKeys, vbYellow, " CD-Keys have been loaded. (", vbWhite, totalExpKeys, vbYellow, " expansion keys)"
  
    calculateAvailableSockets
End Sub

Private Sub lblReloadProxies_Click()
    Dim pl As ProxiesLoaded
    
    pl = loadProxies()
    
    AddChat vbWhite, proxies.countProxies(), vbYellow, " proxies have been loaded."
    
    If (pl.loadedCount > 0) Then
        If (pl.maxProxiesReached) Then
            AddChat vbRed, "Max limit of ", vbWhite, pl.loadedCount, vbRed, " proxies reached!"
        End If
    End If
    
    calculateAvailableSockets
End Sub

Private Sub lblStart_Click()
    If (isTesting) Then
        stopTesting vbYellow, "Testing stopped. click ""start"" to test again."
    Else
        If (Not hasConfig) Then
            MsgBox "Cannot start without a valid config.", vbOKOnly Or vbExclamation, PROGRAM_NAME
            Exit Sub
        End If
  
        If (testedNonExpKeys = totalNonExpKeys) Then
            Dim msg As String
      
            msg = IIf(totalNonExpKeys = 0, "There are no keys available to test", "All non-expansion keys have been tested")
    
            MsgBox msg & ".", vbOKOnly Or vbInformation, PROGRAM_NAME
            Exit Sub
        End If
  
        If (socketsAvailable = 0) Then
            MsgBox "There are no sockets available for testing.", vbOKOnly Or vbInformation, PROGRAM_NAME
            Exit Sub
        End If
  
        If (config.sockets < config.socketsPerProxy) Then
            config.socketsPerProxy = config.sockets
        End If
    
        rtbChat.text = vbNullString
    
        Dim socketsUnavailable As Integer
        socketsUnavailable = 0
    
        For i = 0 To config.sockets - 1
            Dim canUseCDKey As Boolean, canUseProxy As Boolean
            Dim needsCDKey As Boolean, needsProxy As Boolean
      
            canUseCDKey = (BNETData(i).cdKey <> vbNullString)
      
            If (Not canUseCDKey) Then
                needsCDKey = True
                canUseCDKey = canTestRegularKeys()
            End If
      
            canUseProxy = (BNETData(i).proxyIP <> vbNullString)
      
            If (Not canUseProxy) Then
                needsProxy = True
                canUseProxy = proxies.canAcquireProxy()
            End If
      
            If (canUseCDKey And canUseProxy) Then
                If (needsCDKey) Then
                    Dim fk As FoundKey
        
                    fk = getCDKeyFromList()
          
                    With BNETData(i)
                        .cdKey = fk.cdKey
                        .product = fk.product
                        .productRegular = fk.product
                        .cdKeyIndex = fk.keyIndex
                    End With
                End If

                If (needsProxy) Then
                    Dim pType As clsProxyType
                    Set pType = proxies.getProxy()
          
                    With BNETData(i)
                        .proxyIP = pType.getIP()
                        .proxyPort = pType.getPort()
                        .proxyVersion = pType.getVersion()
                        .proxyIndex = pType.getIndex()
                    End With
                End If
        
                Dim IP As String, port As Long
            
                IP = BNETData(i).proxyIP
                port = BNETData(i).proxyPort
    
                AddChat vbYellow, "Socket #" & i & ": Attempting to connect to " & IP & ":" & port & "."
                tmrCheckFailed(i).Enabled = True
                connectSocket i
            Else
                socketsUnavailable = socketsUnavailable + 1
            End If
      
            canUseCDKey = False
            needsCDKey = False
      
            canUseProxy = False
            needsProxy = False
        Next i
    
        isTesting = True
        lblStart.Caption = "Stop"
    
        If (socketsUnavailable > 0) Then
            AddChat vbRed, socketsUnavailable & " sockets were unavailable for use."
        End If
    
        lblReloadProxies.Enabled = False
        lblReloadCDKeys.Enabled = False
        lblConfig.Enabled = False
    
        hasTestedThisSession = True
        tmrBenchmark.Enabled = True
    End If
End Sub

Private Sub lblStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub lblUpdateLabel_Click()
    If (sckCheckUpdate.State = sckClosed) Then
        sckCheckUpdate.Connect "files.codespeak.org", 80
        manualUpdateCheck = True
    End If
End Sub

Private Sub lblUpdateLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub pbMinimize_Click()
    minimize_to_tray
End Sub

Private Sub pbMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub pbQuit_Click()
    EndAll
End Sub

Private Sub pbQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub rtbChat_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(keyCode, Shift)
End Sub

Private Sub sckBNLS_Connect()
    SendBNLS0x10
End Sub

Private Sub sckBNLS_DataArrival(ByVal bytesTotal As Long)
    Dim data As String, pID As Byte, pLen As Long
  
    sckBNLS.GetData data
  
    CopyMemory pLen, ByVal Mid$(data, 1, 2), 2
    pID = Asc(Mid$(data, 3, 1))
  
    bnlsPacket.SetData Mid$(data, 4)
  
    Select Case pID
        Case &H10: RecvBNLS0x10
    End Select
End Sub

Private Sub sckBNCS_Connect(Index As Integer)
    Select Case BNETData(Index).proxyVersion
        Case "SOCKS4":  sckBNCS(Index).SendData Chr$(&H4) & Chr$(&H1) & Chr$(&H17) & Chr$(&HE0) & P_split(LCase$(config.serverIP)) & vbNullString & Chr$(&H0)
        'Case "SOCKS5": 'sckBNCS(Index).SendData Chr$(&H5) & Chr$(&H0)
        Case "HTTP":    sckBNCS(Index).SendData "CONNECT " & config.serverIP & ":6112 HTTP/1.1" & vbCrLf & vbCrLf
  End Select
End Sub

Private Sub sckBNCS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim data As String, pID As Byte, pLen As Long

    sckBNCS(Index).GetData data
    If (IsProxyPacket(Index, data)) Then Exit Sub
  
    If (Asc(left$(data, 1)) <> &HFF) Then
        assumeSocketDead Index
        Exit Sub
    End If
  
    Do While (Len(data) > 3)
        pID = Asc(Mid$(data, 2, 1))
        CopyMemory pLen, ByVal Mid$(data, 3, 2), 2
        If (pLen < 4) Then Exit Sub
  
        packet(Index).SetData Mid$(data, 5, pLen)

        Select Case pID
            Case &H25:  Recv0x25 Index   'Ping
            Case &H50:  Recv0x50 Index   'Auth info
            Case &H51:  Recv0x51 Index   'Auth check
            Case &H52:  Recv0x52 Index   'Account creation result
            Case &H53:  Recv0x53 Index   'NLS Auth account logon
            Case &H54:  Recv0x54 Index   'NLS Auth account logon proof
            Case &H3A:  Recv0x3A Index   'Logon response 2
            Case &H3D:  Recv0x3D Index   'Create Account 2
            Case &H46:  Recv0x46 Index   'News packet
            Case &HA:   Recv0x0A Index   'Enter chat
        End Select
  
        data = Mid$(data, pLen + 1)
    Loop
End Sub

Private Sub sckBNCS_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddChat vbRed, "Socket #" & Index & " error #" & Number & ": " & Description & "."
  
    Call assumeSocketDead(Index)
End Sub

Private Sub sckCheckUpdate_Connect()
    sckCheckUpdate.SendData "GET /projects/maelstrom/CurrentVersion.txt HTTP/1.1" & vbCrLf _
                          & "User-Agent: Maelstrom/" & PROGRAM_VERSION & vbCrLf _
                          & "Host: files.codespeak.org" & vbCrLf & vbCrLf
End Sub

Private Sub sckCheckUpdate_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    sckCheckUpdate.GetData data
    
    updateString = updateString & data
    tmrCheckUpdate.Enabled = True
End Sub

Private Sub sckCheckUpdate_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Unable to check for update!", vbOKOnly Or vbExclamation, PROGRAM_NAME
    tmrCheckUpdate.Enabled = False
    sckCheckUpdate.Close
End Sub

Private Sub tmrBenchmark_Timer()
    curSeconds = curSeconds + 1
  
    lblControl(TIME_ELAPSED).Caption = returnProperTimeString(curSeconds)
    lblControl(KEYS_PER_SECOND).Caption = Format$(CDbl(Round((testedNonExpKeys + testedExpKeys) / curSeconds, 3)), "0.000")
End Sub

Private Sub tmrCheckBNLS_Timer()
    tmrCheckBNLS.Enabled = False
  
    AddChatB vbRed, "Unable to update version byte for " & requestProduct & "!"
    AddChatB vbRed, "The BNLS server has timed out."
    stopTesting vbYellow, "Check your BNLS server and try again."
End Sub

Private Sub tmrCheckFailed_Timer(Index As Integer)
    Call assumeSocketDead(Index)
End Sub

Public Sub assumeSocketDead(Index As Integer)
    tmrCheckFailed(Index).Enabled = False
    closeSocket Index
      
    AddChat vbRed, "Socket #" & Index & ": Connection to " & BNETData(Index).proxyIP & ":" & BNETData(Index).proxyPort & " failed!"
      
    If (config.skipFailedProxies) Then
        proxies.onProxyFail BNETData(Index).proxyIndex
    End If
      
    BNETData(Index).proxyIP = vbNullString
    BNETData(Index).proxyPort = 0
    BNETData(Index).proxyVersion = vbNullString
    BNETData(Index).proxyIndex = 0
    BNETData(Index).numTested = 0
    
    If (proxies.canAcquireProxy()) Then
        Dim IP As String, port As Long, version As String, proxyIndex As Long, pType As clsProxyType
        Set pType = proxies.getProxy()
    
        With pType
            IP = .getIP()
            port = .getPort()
            version = .getVersion()
            proxyIndex = .getIndex()
        End With
    
        BNETData(Index).proxyIP = IP
        BNETData(Index).proxyPort = port
        BNETData(Index).proxyVersion = version
        BNETData(Index).proxyIndex = proxyIndex

        AddChat vbYellow, "Socket #" & Index & ": Attempting to connect to " & IP & ":" & port & "."
    
        connectSocket Index
        tmrCheckFailed(Index).Enabled = True
    Else
        AddChat vbRed, "Socket #" & Index & ": The proxy list has run out."
        markSocketDead Index
  
        If (socketsAvailable = 0) Then
            AddChat vbYellow, "All proxies have been used up."
            lblStart_Click
        End If
    End If
End Sub

Private Sub tmrReconnect_Timer(Index As Integer)
    Dim attemptConnection As Boolean, reportConnection As Boolean
  
    tmrReconnect(Index).Enabled = False
  
    If (config.testCountPerProxy > 0 And BNETData(Index).numTested = config.testCountPerProxy) Then
        AddChat vbYellow, "Socket #" & Index & ": Max tests for this proxy reached."
    
        BNETData(Index).numTested = 0
        proxies.decrementProxyUse BNETData(Index).proxyIndex
    
        If (proxies.canAcquireProxy()) Then
            Dim IP As String, port As Long, version As String, proxyIndex As Long
            Dim pType As clsProxyType
            Set pType = proxies.getProxy()
      
            With pType
                IP = .getIP()
                port = .getPort()
                version = .getVersion()
                proxyIndex = .getIndex()
            End With
    
            BNETData(Index).proxyIP = IP
            BNETData(Index).proxyPort = port
            BNETData(Index).proxyVersion = version
            BNETData(Index).proxyIndex = proxyIndex
    
            attemptConnection = True
            reportConnection = True
        Else
            closeSocket Index
            AddChat vbRed, "Socket #" & Index & ": The proxy list has run out."
            markSocketDead Index
  
            If (socketsAvailable = 0) Then
                AddChat vbYellow, "All connections have been attempted."
                lblStart_Click
                Exit Sub
            End If
        End If
    Else
        IP = BNETData(Index).proxyIP
        port = BNETData(Index).proxyPort
  
        attemptConnection = True
    End If

    If (attemptConnection) Then
        If (reportConnection) Then
            AddChat vbYellow, "Socket #" & Index & ": Attempting to connect to " & IP & ":" & port & "."
        End If
    
        connectSocket Index
        tmrCheckFailed(Index).Enabled = True
    End If
End Sub

Public Sub checkForQuitShortcut(key As Integer, Shift As Integer)
    If (key = 115 And Shift = 4) Then EndAll
End Sub

Private Sub tmrCheckUpdate_Timer()
    On Error GoTo err
  
    Dim versionToCheck As String, updateMsg As String, msgBoxResult As Integer
  
    tmrCheckUpdate.Enabled = False
    versionToCheck = Split(updateString, "Content-Type: text/plain" & vbCrLf & vbCrLf)(1)
    
    If (isNewVersion(versionToCheck)) Then
        updateMsg = "There is a new update for Maelstrom!" & vbNewLine & vbNewLine & "Your version: " & PROGRAM_VERSION & " new version: " & versionToCheck & vbNewLine & vbNewLine _
                  & "Would you like to visit the downloads page for updates?"
    
        msgBoxResult = MsgBox(updateMsg, vbYesNo Or vbInformation, "New Maelstrom version available!")

        If (msgBoxResult = vbYes) Then
            ShellExecute 0, "open", RELEASES_URL, vbNullString, vbNullString, 4
        End If
    Else
        If (manualUpdateCheck) Then
            MsgBox "There is no new version at this time.", vbOKOnly Or vbInformation, PROGRAM_NAME
            manualUpdateCheck = False
        End If
    End If
  
err:
    If err.Number > 0 Then
        err.Clear
        MsgBox "Unable to check for update!", vbOKOnly Or vbExclamation, PROGRAM_NAME
    End If

    updateString = vbNullString
    sckCheckUpdate.Close
End Sub

Private Sub tmrWaitLoad_Timer()
    tmrWaitLoad.Enabled = False
  
    initializeGatewayList
  
    If (dicGatewayIPs.count) = 0 Then
        MsgBox "Unable to contact the Battle.Net servers. Maelstrom cannot continue.", vbOKOnly Or vbCritical, PROGRAM_NAME
        EndAll
  
        Exit Sub
    End If
  
    bnlsPacket.setDetails sckBNLS, PacketType.BNLS
    setupHashFiles

    If (Dir$(App.path & "\Config.ini") = vbNullString) Then
        lblStart.Enabled = True
        makeDefaultValues
    
        MsgBox "No default config found. A default config will be created for you." & vbNewLine _
            & "The config window will open with default values.", vbOKOnly Or vbInformation, PROGRAM_NAME

        frmConfig.Show
    Else
        Dim dicErrors As Dictionary
        Set dicErrors = loadConfig()

        If (dicErrors.count > 0) Then
            lblStart.Enabled = False
    
            MsgBox "There were issues while loading the configuration. The config" & vbNewLine _
                & "window will now be opened so the issues can be corrected.", vbOKOnly Or vbExclamation, PROGRAM_NAME
  
            frmConfig.markErrorLocations dicErrors
            frmConfig.Show
        Else
            hasConfig = True
        End If
    End If

    Dim pl As ProxiesLoaded

    pl = loadProxies()
    loadCDKeys
  
    AddChat vbWhite, proxies.countProxies(), vbYellow, " proxies have been loaded."
    AddChat vbWhite, config.sockets, vbYellow, " sockets have been loaded. (", vbWhite, config.socketsPerProxy, vbYellow, " per proxy)"
    AddChat vbWhite, totalNonExpKeys, vbYellow, " CD-Keys have been loaded."

    If (totalExpKeys > 0) Then
         AddChat vbWhite, totalExpKeys, vbYellow, " expansion CD-Keys have been loaded."
    End If

    If (pl.loadedCount > 0) Then
        If (pl.maxProxiesReached) Then
            AddChat vbRed, "Max limit of ", vbWhite, pl.loadedCount, vbRed, " proxies reached!"
        End If
    End If

    If (config.sockets > 0) Then
        setupConnectionData config.sockets
    End If

    calculateAvailableSockets
  
    If (config.saveWindowPosition) Then
        tempValue = ReadINI("Window", "Top", "Config.ini")
    
        If (IsNumericB(tempValue)) Then
            Me.top = tempValue
        End If
    
        tempValue = ReadINI("Window", "Left", "Config.ini")
    
        If (IsNumericB(tempValue)) Then
            Me.left = tempValue
        End If
    End If
    
    If (config.cdKeyProfile <> vbNullString) Then
        Dim fullProfileName As String
        fullProfileName = config.cdKeyProfile

        If (config.addRealmToProfile) Then
            fullProfileName = fullProfileName & " @ " & config.ServerRealm
        End If

        lblControl(CDKEY_PROFILE).Caption = "Using CD-Key Profile: " & fullProfileName
    Else
        lblControl(CDKEY_PROFILE).Caption = "CD-Key Profile Not Configured"
    End If
    
    If (config.checkUpdateOnStartup) Then
        If (sckCheckUpdate.State = sckClosed) Then
            sckCheckUpdate.Connect "files.codespeak.org", 80
        End If
    End If
End Sub
 
