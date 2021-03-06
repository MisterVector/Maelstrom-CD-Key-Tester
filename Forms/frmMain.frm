VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H007F7F7F&
   BorderStyle     =   0  'None
   Caption         =   " Maelstrom CD-Key Tester"
   ClientHeight    =   6285
   ClientLeft      =   8805
   ClientTop       =   3510
   ClientWidth     =   11610
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   11610
   Begin VB.Timer tmrWaitLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   5040
   End
   Begin VB.Timer tmrCheckBNLS 
      Enabled         =   0   'False
      Left            =   240
      Top             =   5520
   End
   Begin VB.Timer tmrBenchmark 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   4560
   End
   Begin MSWinsockLib.Winsock sckBNLS 
      Left            =   720
      Top             =   4560
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
      TabIndex        =   49
      Top             =   -15
      Width           =   480
   End
   Begin MSWinsockLib.Winsock sckBNCS 
      Index           =   0
      Left            =   240
      Top             =   4560
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
      TabIndex        =   19
      Top             =   -15
      Width           =   480
   End
   Begin VB.Timer tmrCheckFailed 
      Enabled         =   0   'False
      Index           =   0
      Left            =   1200
      Top             =   5040
   End
   Begin VB.Timer tmrReconnect 
      Enabled         =   0   'False
      Index           =   0
      Left            =   240
      Top             =   5040
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   3135
      Left            =   120
      TabIndex        =   18
      Top             =   3000
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
      TabIndex        =   67
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
      TabIndex        =   66
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
      TabIndex        =   65
      Top             =   5280
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
      TabIndex        =   64
      Top             =   5280
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
      TabIndex        =   63
      Top             =   5640
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
      TabIndex        =   62
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   7320
      X2              =   11400
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   11400
      X2              =   7320
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   7320
      X2              =   11400
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   7320
      X2              =   11400
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   7320
      X2              =   11400
      Y1              =   4080
      Y2              =   4080
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
      Left            =   10200
      TabIndex        =   61
      Top             =   4200
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
      TabIndex        =   60
      Top             =   4200
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
      TabIndex        =   59
      Top             =   4200
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
      TabIndex        =   58
      Top             =   4200
      Width           =   855
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
      TabIndex        =   57
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblControl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0:00:00:00"
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
      Left            =   9600
      TabIndex        =   56
      Top             =   4560
      Width           =   1815
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
      TabIndex        =   55
      Top             =   4920
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
      TabIndex        =   54
      Top             =   4560
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
      Index           =   96
      Left            =   10560
      TabIndex        =   53
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
      TabIndex        =   52
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
      TabIndex        =   51
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
      TabIndex        =   50
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
      Index           =   42
      Left            =   8640
      TabIndex        =   48
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
      TabIndex        =   47
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
      TabIndex        =   46
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
      Index           =   51
      Left            =   8640
      TabIndex        =   45
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   48
      Left            =   5880
      TabIndex        =   43
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
      TabIndex        =   42
      Top             =   2160
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
      TabIndex        =   41
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
      TabIndex        =   39
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
      TabIndex        =   38
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
      Index           =   32
      Left            =   7680
      TabIndex        =   37
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
      Index           =   43
      Left            =   9465
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      Index           =   41
      Left            =   7665
      TabIndex        =   33
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      Index           =   47
      Left            =   5040
      TabIndex        =   30
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
      TabIndex        =   29
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
      Index           =   29
      Left            =   5025
      TabIndex        =   28
      Top             =   1800
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      Index           =   46
      Left            =   4080
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   2160
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
      TabIndex        =   40
      Top             =   3840
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
      TabIndex        =   22
      Top             =   5640
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
      TabIndex        =   21
      Top             =   3000
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
      TabIndex        =   20
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
      TabIndex        =   15
      Top             =   3480
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
      TabIndex        =   14
      Top             =   3480
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
      TabIndex        =   13
      Top             =   3480
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
      TabIndex        =   12
      Top             =   3480
      Width           =   1815
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
    nid.szTip = PROGRAM_TITLE & vbNullChar
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(keyCode, Shift)
End Sub

Private Sub Form_Load()
    Dim top As Long, left As Long, tempValue As String

    AddChat vbYellow, "Welcome to " & PROGRAM_NAME & " ", vbWhite, "v" & PROGRAM_VERSION, vbYellow, " by Vector."
    lblControl(1).Caption = PROGRAM_TITLE
  
    If (InStr(Command, "--csds-launch") > 0) Then
        loadedFromCSDSClient = True
    End If
  
    tmrWaitLoad.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim msg As Long
    Dim sFilter As String
    msg = x / Screen.TwipsPerPixelX
  
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

Private Sub lblConfig_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub lblControl_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
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

Private Sub lblReloadCDKeys_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub lblReloadProxies_Click()
    Dim pl As ProxiesLoaded
    
    For i = 0 To UBound(BNETData)
        With BNETData(i)
            .proxyIP = vbNullString
            .proxyPort = 0
            .proxyVersion = vbNullString
        End With
    Next i
    
    pl = loadProxies()
    
    AddChat vbWhite, proxies.countProxies(), vbYellow, " proxies have been loaded."
    
    If (pl.loadedCount > 0) Then
        If (pl.maxProxiesReached) Then
            AddChat vbRed, "Max limit of ", vbWhite, pl.loadedCount, vbRed, " proxies reached!"
        End If
    End If
    
    calculateAvailableSockets
End Sub

Private Sub lblReloadProxies_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub lblStart_Click()
    If (isTesting) Then
        stopTesting vbYellow, "Testing stopped. Click ", vbWhite, "Start", vbYellow, " to test again."
    Else
        If (Not hasConfig) Then
            MsgBox "Cannot start without a valid config.", vbOKOnly Or vbExclamation, PROGRAM_TITLE
            Exit Sub
        End If
  
        If (testedNonExpKeys = totalNonExpKeys) Then
            Dim msg As String
      
            msg = IIf(totalNonExpKeys = 0, "There are no keys available to test", "All non-expansion keys have been tested")
    
            MsgBox msg & ".", vbOKOnly Or vbInformation, PROGRAM_TITLE
            Exit Sub
        End If
  
        If (socketsAvailable = 0) Then
            MsgBox "There are no sockets available for testing.", vbOKOnly Or vbInformation, PROGRAM_TITLE
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
      
            canUseCDKey = (BNETData(i).CDKey <> vbNullString)
      
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
                        .CDKey = fk.CDKey
                        .Product = fk.Product
                        .productRegular = fk.Product
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
        
                Dim IP As String, Port As Long
            
                IP = BNETData(i).proxyIP
                Port = BNETData(i).proxyPort
    
                AddChat vbYellow, "Socket #" & i & ": Attempting to connect to ", vbWhite, IP & ":" & Port, vbYellow, "."
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
            AddChat vbWhite, socketsUnavailable, vbRed, " sockets were unavailable for use."
        End If
    
        lblReloadProxies.Enabled = False
        lblReloadCDKeys.Enabled = False
        lblConfig.Enabled = False
    
        hasTestedThisSession = True
        tmrBenchmark.Enabled = True
    End If
End Sub

Private Sub lblStart_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub lblUpdateLabel_Click()
    If (loadedFromCSDSClient) Then
        MsgBox "Cannot check update as Maelstrom was loaded by the Code Speak Distribution Client!", vbOKOnly Or vbExclamation, PROGRAM_TITLE
        Exit Sub
    End If

    If (Not checkProgramUpdate(True)) Then
        MsgBox "Unable to check for update!", vbOKOnly Or vbExclamation, PROGRAM_TITLE
    End If
End Sub

Private Sub lblUpdateLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub pbMinimize_Click()
    minimize_to_tray
End Sub

Private Sub pbMinimize_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub pbQuit_Click()
    EndAll
End Sub

Private Sub pbQuit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub rtbChat_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(keyCode, Shift)
End Sub

Private Sub sckBNLS_Connect()
    SendBNLS0x10
End Sub

Private Sub sckBNLS_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String, pID As Byte, pLen As Long
  
    sckBNLS.GetData Data
  
    CopyMemory pLen, ByVal Mid$(Data, 1, 2), 2
    pID = Asc(Mid$(Data, 3, 1))
  
    bnlsPacket.SetData Mid$(Data, 4)
  
    Select Case pID
        Case &H10: RecvBNLS0x10
    End Select
End Sub

Private Sub sckBNCS_Connect(index As Integer)
    Select Case BNETData(index).proxyVersion
        Case "SOCKS4":  sckBNCS(index).SendData Chr$(&H4) & Chr$(&H1) & Chr$(&H17) & Chr$(&HE0) & P_split(LCase$(config.serverIP)) & vbNullString & Chr$(&H0)
        Case "SOCKS5":  sckBNCS(index).SendData Chr$(&H5) & Chr$(&H1) & Chr$(&H0)
        Case "HTTP":    sckBNCS(index).SendData "CONNECT " & config.serverIP & ":6112 HTTP/1.1" & vbCrLf & vbCrLf
  End Select
End Sub

Private Sub sckBNCS_DataArrival(index As Integer, ByVal bytesTotal As Long)
    Dim Data As String, pID As Byte, pLen As Long

    sckBNCS(index).GetData Data
    If (IsProxyPacket(index, Data)) Then Exit Sub
  
    If (Asc(left$(Data, 1)) <> &HFF) Then
        assumeSocketDead index
        Exit Sub
    End If
  
    Do While (Len(Data) > 3)
        pID = Asc(Mid$(Data, 2, 1))
        CopyMemory pLen, ByVal Mid$(Data, 3, 2), 2
        If (pLen < 4) Then Exit Sub
  
        packet(index).SetData Mid$(Data, 5, pLen)

        Select Case pID
            Case &H25:  Recv0x25 index   'Ping
            Case &H50:  Recv0x50 index   'Auth info
            Case &H51:  Recv0x51 index   'Auth check
            Case &H3A:  Recv0x3A index   'Logon response 2
            Case &H3D:  Recv0x3D index   'Create Account 2
            Case &H46:  Recv0x46 index   'News packet
            Case &HA:   Recv0x0A index   'Enter chat
        End Select
  
        Data = Mid$(Data, pLen + 1)
    Loop
End Sub

Private Sub sckBNCS_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddChat vbRed, "Socket #" & index & " error #" & Number & ": " & Description & "."
  
    Call assumeSocketDead(index)
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

Private Sub tmrCheckFailed_Timer(index As Integer)
    Call assumeSocketDead(index)
End Sub

Public Sub assumeSocketDead(index As Integer)
    tmrCheckFailed(index).Enabled = False
    closeSocket index
      
    AddChat vbRed, "Socket #" & index & ": Connection to ", vbWhite, BNETData(index).proxyIP & ":" & BNETData(index).proxyPort, vbRed, " failed!"
      
    If (config.skipFailedProxies) Then
        proxies.onProxyFail BNETData(index).proxyIndex
    End If
      
    BNETData(index).proxyIP = vbNullString
    BNETData(index).proxyPort = 0
    BNETData(index).proxyVersion = vbNullString
    BNETData(index).proxyIndex = 0
    BNETData(index).numTested = 0
    BNETData(index).acceptedAuth = False
    
    If (proxies.canAcquireProxy()) Then
        Dim IP As String, Port As Long, Version As String, proxyIndex As Long, pType As clsProxyType
        Set pType = proxies.getProxy()
    
        With pType
            IP = .getIP()
            Port = .getPort()
            Version = .getVersion()
            proxyIndex = .getIndex()
        End With
    
        BNETData(index).proxyIP = IP
        BNETData(index).proxyPort = Port
        BNETData(index).proxyVersion = Version
        BNETData(index).proxyIndex = proxyIndex

        AddChat vbYellow, "Socket #" & index & ": Attempting to connect to ", vbWhite, IP & ":" & Port, vbYellow, "."
    
        connectSocket index
        tmrCheckFailed(index).Enabled = True
    Else
        AddChat vbRed, "Socket #" & index & ": The proxy list has run out."
        markSocketDead index
  
        If (socketsAvailable = 0) Then
            AddChat vbYellow, "All proxies have been used up."
            lblStart_Click
        End If
    End If
End Sub

Private Sub tmrReconnect_Timer(index As Integer)
    Dim attemptConnection As Boolean, reportConnection As Boolean
  
    tmrReconnect(index).Enabled = False
  
    If (config.testCountPerProxy > 0 And BNETData(index).numTested = config.testCountPerProxy) Then
        AddChat vbYellow, "Socket #" & index & ": Max tests for this proxy reached."
    
        BNETData(index).numTested = 0
        proxies.decrementProxyUse BNETData(index).proxyIndex
    
        If (proxies.canAcquireProxy()) Then
            Dim IP As String, Port As Long, Version As String, proxyIndex As Long
            Dim pType As clsProxyType
            Set pType = proxies.getProxy()
      
            With pType
                IP = .getIP()
                Port = .getPort()
                Version = .getVersion()
                proxyIndex = .getIndex()
            End With
    
            BNETData(index).proxyIP = IP
            BNETData(index).proxyPort = Port
            BNETData(index).proxyVersion = Version
            BNETData(index).proxyIndex = proxyIndex
    
            attemptConnection = True
            reportConnection = True
        Else
            closeSocket index
            AddChat vbRed, "Socket #" & index & ": The proxy list has run out."
            markSocketDead index
  
            If (socketsAvailable = 0) Then
                AddChat vbYellow, "All connections have been attempted."
                lblStart_Click
                Exit Sub
            End If
        End If
    Else
        IP = BNETData(index).proxyIP
        Port = BNETData(index).proxyPort
  
        attemptConnection = True
    End If

    If (attemptConnection) Then
        If (reportConnection) Then
            AddChat vbYellow, "Socket #" & index & ": Attempting to connect to ", vbWhite, IP & ":" & Port, vbYellow, "."
        End If
    
        connectSocket index
        tmrCheckFailed(index).Enabled = True
    End If
End Sub

Public Sub checkForQuitShortcut(key As Integer, Shift As Integer)
    If (key = 115 And Shift = 4) Then EndAll
End Sub

Private Sub tmrWaitLoad_Timer()
    tmrWaitLoad.Enabled = False
  
    initializeGatewayList
  
    If (dicGatewayIPs.count) = 0 Then
        MsgBox "Unable to contact the Battle.Net servers. Maelstrom cannot continue.", vbOKOnly Or vbCritical, PROGRAM_TITLE
        EndAll
  
        Exit Sub
    End If
  
    bnlsPacket.setDetails sckBNLS, PacketType.BNLS
    setupHashFiles

    If (Dir$(App.path & "\Config.ini") = vbNullString) Then
        lblStart.Enabled = True
        makeDefaultValues
    
        MsgBox "No default config found. A default config will be created for you." & vbNewLine _
            & "The config window will open with default values.", vbOKOnly Or vbInformation, PROGRAM_TITLE

        frmConfig.Show
    Else
        Dim dicErrors As Dictionary
        Set dicErrors = loadConfig()

        If (dicErrors.count > 0) Then
            lblStart.Enabled = False
    
            MsgBox "There were issues while loading the configuration. The config" & vbNewLine _
                & "window will now be opened so the issues can be corrected.", vbOKOnly Or vbExclamation, PROGRAM_TITLE
  
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
    
    If (config.checkUpdateOnStartup And Not loadedFromCSDSClient) Then
        If (Not checkProgramUpdate(False)) Then
            MsgBox "Unable to check for update!", vbOKOnly Or vbExclamation, PROGRAM_TITLE
        End If
    End If
End Sub
 
