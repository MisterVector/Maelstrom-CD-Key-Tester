VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSaveWindowPosition 
      BackColor       =   &H00404040&
      Caption         =   "Save Window Position"
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
      Height          =   375
      Left            =   240
      TabIndex        =   46
      Top             =   6360
      Width           =   2655
   End
   Begin VB.CheckBox chkAddRealmToProfile 
      BackColor       =   &H00404040&
      Caption         =   "Add Realm To Profile"
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
      Height          =   330
      Left            =   3960
      TabIndex        =   21
      Top             =   6030
      Width           =   2580
   End
   Begin VB.CheckBox chkSkipFailedProxies 
      BackColor       =   &H00404040&
      Caption         =   "Skip Failed Proxies"
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
      Left            =   240
      TabIndex        =   19
      Top             =   5640
      Width           =   2355
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   16
      Left            =   5640
      TabIndex        =   17
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CheckBox chkSaveGoodProxies 
      BackColor       =   &H00404040&
      Caption         =   "Save Good Proxies"
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
      Height          =   345
      Left            =   240
      TabIndex        =   20
      Top             =   6000
      Width           =   2355
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   15
      Left            =   5160
      TabIndex        =   16
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   14
      Left            =   6720
      TabIndex        =   15
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   13
      Left            =   5160
      TabIndex        =   14
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   11
      Left            =   5760
      TabIndex        =   12
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   10
      Left            =   6360
      TabIndex        =   9
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   9
      Left            =   6360
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   8
      Left            =   3120
      TabIndex        =   11
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   7
      Left            =   3120
      TabIndex        =   13
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   6
      Left            =   3240
      TabIndex        =   8
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   7
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   10
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.ComboBox cmbServer 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "frmConfig.frx":0000
      Left            =   1080
      List            =   "frmConfig.frx":0002
      TabIndex        =   5
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CheckBox chkAddDateToTested 
      BackColor       =   &H00404040&
      Caption         =   "Add Date To Tested"
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
      Left            =   240
      TabIndex        =   18
      Top             =   5280
      Width           =   2445
   End
   Begin VB.PictureBox pbQuit 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   6720
      Picture         =   "frmConfig.frx":0004
      ScaleHeight     =   450
      ScaleWidth      =   465
      TabIndex        =   22
      Top             =   0
      Width           =   465
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "CD-Key Profile:"
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
      Index           =   21
      Left            =   3960
      TabIndex        =   45
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label lblRestoreDefaults 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Restore Defaults"
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
      Height          =   375
      Left            =   2040
      TabIndex        =   44
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label lblConfig 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version Bytes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   24
      Left            =   3960
      TabIndex        =   43
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Height          =   495
      Left            =   240
      TabIndex        =   42
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CANCEL"
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
      Height          =   495
      Left            =   5280
      TabIndex        =   41
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Warcraft II::"
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
      Index           =   18
      Left            =   3960
      TabIndex        =   40
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Diablo II::"
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
      Index           =   17
      Left            =   5640
      TabIndex        =   39
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Warcraft III:"
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
      Index           =   16
      Left            =   3960
      TabIndex        =   38
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "VerByte Server:"
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
      Index           =   15
      Left            =   3960
      TabIndex        =   37
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Reconnect Time:"
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
      Index           =   14
      Left            =   3960
      TabIndex        =   36
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Timeout Time:"
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
      Index           =   13
      Left            =   3960
      TabIndex        =   35
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Count Per Proxy:"
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
      Index           =   12
      Left            =   240
      TabIndex        =   34
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Expansion Tests Per Regular Key:"
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
      Height          =   495
      Index           =   11
      Left            =   240
      TabIndex        =   33
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Per Proxy:"
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
      Left            =   2040
      TabIndex        =   32
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Sockets:"
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
      Index           =   9
      Left            =   240
      TabIndex        =   31
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Channel To Join:"
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
      Index           =   8
      Left            =   240
      TabIndex        =   30
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
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
      Left            =   240
      TabIndex        =   29
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblConfig 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warcraft III Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   28
      Top             =   1920
      Width           =   6855
   End
   Begin VB.Label lblConfig 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Non-Warcraft III Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   27
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Index           =   1
      Left            =   3960
      TabIndex        =   26
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   3960
      TabIndex        =   25
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
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
      Left            =   240
      TabIndex        =   24
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
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
      Left            =   240
      TabIndex        =   23
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblConfig 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Maelstrom Configuration"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const CONFIG_NUMERIC_INDEXES = "5 6 7 8 9 10"
Private Const CONFIG_HEX_INDEXES = "13 14 15"
Private Const CONFIG_DEFAULT_TEXTBOX_IDXS = "5 6 7 8 9 10 11 13 14 15"

Private Sub chkAddDateToTested_Click()
  If (chkAddDateToTested.BackColor <> FRM_BACK_COLOR) Then
    chkAddDateToTested.BackColor = FRM_BACK_COLOR
  End If
End Sub

Private Sub chkAddDateToTested_KeyDown(keyCode As Integer, shift As Integer)
  Call checkForQuitShortcut(Me, keyCode, shift)
End Sub

Private Sub chkAddDateToTested_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub chkAddRealmToProfile_Click()
  If (chkAddRealmToProfile.BackColor <> FRM_BACK_COLOR) Then
    chkAddRealmToProfile.BackColor = FRM_BACK_COLOR
  End If
End Sub

Private Sub chkAddRealmToProfile_KeyDown(keyCode As Integer, shift As Integer)
  Call checkForQuitShortcut(Me, keyCode, shift)
End Sub

Private Sub chkAddRealmToProfile_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub chkSaveGoodProxies_KeyDown(keyCode As Integer, shift As Integer)
  Call checkForQuitShortcut(Me, keyCode, shift)
End Sub

Private Sub chkSaveGoodProxies_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub chkSaveGoodProxies_Click()
  If (chkSaveGoodProxies.BackColor <> FRM_BACK_COLOR) Then
    chkSaveGoodProxies.BackColor = FRM_BACK_COLOR
  End If
End Sub

Private Sub chkSkipFailedProxies_KeyDown(keyCode As Integer, shift As Integer)
  Call checkForQuitShortcut(Me, keyCode, shift)
End Sub

Private Sub chkSkipFailedProxies_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub chkSkipFailedProxies_Click()
  If (chkSkipFailedProxies.BackColor <> FRM_BACK_COLOR) Then
    chkSkipFailedProxies.BackColor = FRM_BACK_COLOR
  End If
  
  chkSaveGoodProxies.Enabled = IIf(chkSkipFailedProxies.value = 1, True, False)
End Sub

Private Sub chkSaveWindowPosition_KeyDown(keyCode As Integer, shift As Integer)
  Call checkForQuitShortcut(Me, keyCode, shift)
End Sub

Private Sub chkSaveWindowPosition_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub chkSaveWindowPosition_Click()
  If (chkSaveWindowPosition.BackColor <> FRM_BACK_COLOR) Then
    chkSaveWindowPosition.BackColor = FRM_BACK_COLOR
  End If
End Sub

Private Sub cmbServer_Change()
  If (cmbServer.BackColor <> TXT_BACK_COLOR) Then
    cmbServer.BackColor = TXT_BACK_COLOR
  End If
End Sub

Private Sub cmbServer_Click()
  If (cmbServer.BackColor <> TXT_BACK_COLOR) Then
    cmbServer.BackColor = TXT_BACK_COLOR
  End If
End Sub

Private Sub cmbServer_KeyDown(keyCode As Integer, shift As Integer)
  Call checkForQuitShortcut(Me, keyCode, shift)
End Sub

Private Sub Form_KeyDown(keyCode As Integer, shift As Integer)
  Call checkForQuitShortcut(Me, keyCode, shift)
End Sub

Private Sub Form_Load()
  Dim gateway As Variant

  txtConfig(CONFIG_USERNAME).text = config.name
  txtConfig(CONFIG_PASSWORD).text = config.password
  
  txtConfig(CONFIG_USERNAMEW3).text = config.nameW3
  txtConfig(CONFIG_PASSWORDW3).text = config.passwordW3
  
  If (config.server <> vbNullString) Then
    cmbServer.text = config.server
    cmbServer.AddItem config.server
  End If
  
  For Each gateway In dicGatewayIPs.Keys
    Dim IP As Variant, IPs As Variant
  
    IPs = dicGatewayIPs.Item(gateway)
    
    cmbServer.AddItem vbNullString
    cmbServer.AddItem gateway
    
    For Each IP In IPs
      cmbServer.AddItem IP
    Next
  Next
  
  txtConfig(CONFIG_HOME_CHANNEL).text = config.homeChannel
  txtConfig(CONFIG_SOCKETS).text = config.sockets
  txtConfig(CONFIG_SOCKETS_PER_PROXY).text = config.socketsPerProxy
  txtConfig(CONFIG_EXP_TESTS_PER_REG_KEY).text = config.expansionTestsPerRegularKey
  txtConfig(CONFIG_TEST_COUNT_PER_PROXY).text = config.testCountPerProxy
  txtConfig(CONFIG_CHECK_FAILURE).text = config.checkFailure
  txtConfig(CONFIG_RECONNECT_TIME).text = config.reconnectTime
  txtConfig(CONFIG_CDKEY_PROFILE).text = config.cdKeyProfile
  
  chkAddDateToTested.value = IIf(config.addDateToTested, 1, 0)
  chkSaveGoodProxies.value = IIf(config.saveGoodProxies, 1, 0)
  chkSkipFailedProxies.value = IIf(config.skipFailedProxies, 1, 0)
  chkAddRealmToProfile.value = IIf(config.addRealmToProfile, 1, 0)
  chkSaveWindowPosition.value = IIf(config.saveWindowPosition, 1, 0)
  
  If (chkSkipFailedProxies.value = 0) Then
    chkSaveGoodProxies.Enabled = False
  End If
  
  txtConfig(CONFIG_BNLS_SERVER).text = config.bnlsServer
  
  txtConfig(CONFIG_VERBYTE_W2BN).text = Hex(config.W2BNVerByte)
  txtConfig(CONFIG_VERBYTE_D2DV).text = Hex(config.D2DVVerByte)
  txtConfig(CONFIG_VERBYTE_WAR3).text = Hex(config.WAR3VerByte)
End Sub

Private Sub Form_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If (isClosing) Then Exit Sub

  If (Not hasConfig) Then
    Dim msgBoxResult As Integer
    msgBoxResult = MsgBox("Are you sure you want to cancel without a valid configuration?", vbYesNo & vbQuestion, PROGRAM_NAME)
  
    If (msgBoxResult = vbNo) Then Cancel = 1
  End If
  
  frmMain.lblStart.Enabled = True
End Sub

Private Sub lblCancel_Click()
  Unload Me
End Sub

Private Sub lblCancel_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub lblConfig_MouseMove(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub lblOk_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub lblOk_Click()
  Dim errors As Integer, oldProfile As String

  errors = markFormErrors()
  
  If (errors > 0) Then Exit Sub
  
  oldProfile = config.cdKeyProfile
  
  If (config.addRealmToProfile) Then
    oldProfile = oldProfile & " @ " & config.ServerRealm
  End If
  
  config.server = cmbServer.text
  config.serverIP = getProperGateway(config.server)
  
  Dim sr As ServerRealm
  
  sr = serverToRealm(config.server)
  
  config.ServerRealm = sr.realm
  config.serverRealmW3 = sr.realmW3
  
  config.name = txtConfig(CONFIG_USERNAME).text
  config.password = txtConfig(CONFIG_PASSWORD).text
  
  config.nameW3 = txtConfig(CONFIG_USERNAMEW3).text
  config.passwordW3 = txtConfig(CONFIG_PASSWORDW3).text
  
  config.homeChannel = txtConfig(CONFIG_HOME_CHANNEL).text
  config.sockets = txtConfig(CONFIG_SOCKETS).text
  config.socketsPerProxy = txtConfig(CONFIG_SOCKETS_PER_PROXY).text
  config.expansionTestsPerRegularKey = txtConfig(CONFIG_EXP_TESTS_PER_REG_KEY).text
  config.testCountPerProxy = txtConfig(CONFIG_TEST_COUNT_PER_PROXY).text
  config.checkFailure = txtConfig(CONFIG_CHECK_FAILURE).text
  config.reconnectTime = txtConfig(CONFIG_RECONNECT_TIME).text
  config.bnlsServer = txtConfig(CONFIG_BNLS_SERVER).text
  config.cdKeyProfile = txtConfig(CONFIG_CDKEY_PROFILE).text
  
  config.W2BNVerByte = "&H" & txtConfig(CONFIG_VERBYTE_W2BN).text
  config.D2DVVerByte = "&H" & txtConfig(CONFIG_VERBYTE_D2DV).text
  config.WAR3VerByte = "&H" & txtConfig(CONFIG_VERBYTE_WAR3).text
  
  config.addDateToTested = IIf(chkAddDateToTested.value, True, False)
  config.saveGoodProxies = IIf(chkSaveGoodProxies.value, True, False)
  config.skipFailedProxies = IIf(chkSkipFailedProxies.value, True, False)
  config.addRealmToProfile = IIf(chkAddRealmToProfile.value, True, False)
  config.saveWindowPosition = IIf(chkSaveWindowPosition.value, True, False)
  
  writeConfig
  
  If (hasTestedThisSession And loadedSockets <> config.sockets) Then
    hasTestedThisSession = False
    
    restoreKeysToList
    proxies.resetProxies
  End If
  
  If (loadedSockets <> config.sockets) Then
    AddChat vbWhite, config.sockets, vbYellow, " sockets have been loaded. (", vbWhite, config.socketsPerProxy, vbYellow, " per proxy)"
    calculateAvailableSockets
  End If
  
  setupConnectionData config.sockets
  
  If (config.cdKeyProfile <> vbNullString) Then
    Dim fullProfileName As String
    fullProfileName = config.cdKeyProfile
  
    If (config.addRealmToProfile) Then
      fullProfileName = fullProfileName & " @ " & config.ServerRealm
    End If
  
    If (oldProfile <> fullProfileName) Then
      AddChat vbYellow, "Using CD-Key profile """ & fullProfileName & """."
    End If
  End If
  
  hasConfig = True
  
  Unload Me
End Sub

Private Sub lblRestoreDefaults_Click()
  txtConfig(CONFIG_SOCKETS).text = DEFAULT_SOCKETS
  txtConfig(CONFIG_SOCKETS_PER_PROXY).text = DEFAULT_SOCKETS_PER_PROXY
  txtConfig(CONFIG_EXP_TESTS_PER_REG_KEY).text = DEFAULT_EXP_TESTS_PER_REG_KEY
  txtConfig(CONFIG_TEST_COUNT_PER_PROXY).text = DEFAULT_TEST_COUNT_PER_PROXY
  txtConfig(CONFIG_CHECK_FAILURE).text = DEFAULT_CHECK_FAILURE
  txtConfig(CONFIG_RECONNECT_TIME).text = DEFAULT_RECONNECT_TIME
  txtConfig(CONFIG_BNLS_SERVER).text = DEFAULT_BNLS_SERVER
  
  txtConfig(CONFIG_VERBYTE_W2BN).text = Hex(DEFAULT_VERBYTE_W2BN)
  txtConfig(CONFIG_VERBYTE_D2DV).text = Hex(DEFAULT_VERBYTE_D2DV)
  txtConfig(CONFIG_VERBYTE_WAR3).text = Hex(DEFAULT_VERBYTE_WAR3)
  
  Dim defaultIndexes() As String
  defaultIndexes = Split(CONFIG_DEFAULT_TEXTBOX_IDXS, " ")
  
  For i = 0 To UBound(defaultIndexes)
    txtConfig(defaultIndexes(i)).BackColor = TXT_BACK_COLOR
  Next i
  
  chkAddDateToTested.value = IIf(DEFAULT_ADD_DATE_TO_TESTED, 1, 0)
  chkSaveGoodProxies.value = IIf(DEFAULT_SAVE_GOOD_PROXIES, 1, 0)
  chkSkipFailedProxies.value = IIf(DEFAULT_SKIP_FAILED_PROXIES, 1, 0)
  chkAddRealmToProfile.value = IIf(DEFAULT_ADD_REALM_TO_PROFILE, 1, 0)
  chkSaveWindowPosition.value = IIf(DEFAULT_SAVE_WINDOW_POSITION, 1, 0)
  
  If (chkAddDateToTested.BackColor <> FRM_BACK_COLOR) Then
    chkAddDateToTested.BackColor = FRM_BACK_COLOR
  End If
  
  If (chkSaveGoodProxies.BackColor <> FRM_BACK_COLOR) Then
    chkSaveGoodProxies.BackColor = FRM_BACK_COLOR
  End If
  
  If (chkSkipFailedProxies.BackColor <> FRM_BACK_COLOR) Then
    chkSkipFailedProxies.BackColor = FRM_BACK_COLOR
  End If
  
  If (chkAddRealmToProfile.BackColor <> FRM_BACK_COLOR) Then
    chkAddRealmToProfile.BackColor = FRM_BACK_COLOR
  End If
  
  If (chkSaveWindowPosition.BackColor <> FRM_BACK_COLOR) Then
    chkSaveWindowPosition.BackColor = FRM_BACK_COLOR
  End If
  
  chkSaveGoodProxies.Enabled = IIf(chkSkipFailedProxies.value = 1, True, False)
End Sub

Private Sub lblRestoreDefaults_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub pbQuit_KeyDown(keyCode As Integer, shift As Integer)
  Call checkForQuitShortcut(Me, keyCode, shift)
End Sub

Private Sub pbQuit_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
  Call moveEntireForm(Me, Button)
End Sub

Private Sub pbQuit_Click()
  Unload Me
End Sub

Private Sub txtConfig_Change(Index As Integer)
  If (txtConfig(Index).BackColor <> TXT_BACK_COLOR) Then
    txtConfig(Index).BackColor = TXT_BACK_COLOR
  End If
End Sub

Private Sub txtConfig_KeyDown(Index As Integer, keyCode As Integer, shift As Integer)
  Call checkForQuitShortcut(Me, keyCode, shift)
End Sub

Private Function markFormErrors() As Integer
  Dim o As Control, errors As Integer
  
  For Each o In Me.Controls
    If (TypeOf o Is TextBox) Then
      Dim t As TextBox, lenRequire As Integer, hasError As Boolean
      hasError = False
      
      Set t = o
      lenRequire = IIf(t.Index = CONFIG_USERNAME Or t.Index = CONFIG_USERNAMEW3 Or t.Index = CONFIG_BNLS_SERVER, 3, IIf(t.Index = CONFIG_CDKEY_PROFILE, 0, 1))

      If (Len(t.text) < lenRequire) Then
        t.BackColor = TXT_ERROR_COLOR
        errors = errors + 1
        hasError = True
      End If

      If (Not hasError) Then
        Dim txtIndex As Integer, numConfig() As String
        
        numConfig = Split(CONFIG_NUMERIC_INDEXES)
        
        For i = 0 To UBound(numConfig)
          txtIndex = numConfig(i)
        
          If (t.Index = txtIndex) Then
            hasError = True
            
            If (IsNumericB(t.text)) Then
              Dim minNumber As Integer, maxNumber As Integer
              
              If (txtIndex = CONFIG_TEST_COUNT_PER_PROXY) Then
                minNumber = 0
              Else
                minNumber = 1
              End If
            
              If (txtIndex = CONFIG_SOCKETS) Then
                maxNumber = MAX_SOCKETS
              ElseIf (txtIndex = CONFIG_SOCKETS_PER_PROXY) Then
                maxNumber = MAX_SOCKETS_PER_PROXY
              ElseIf (txtIndex = CONFIG_EXP_TESTS_PER_REG_KEY) Then
                maxNumber = MAX_EXP_TESTS_PER_REG_KEY
              ElseIf (txtIndex = CONFIG_TEST_COUNT_PER_PROXY) Then
                maxNumber = MAX_TEST_COUNT_PER_PROXY
              ElseIf (txtIndex = CONFIG_RECONNECT_TIME) Then
                maxNumber = MAX_RECONNECT_TIME
              ElseIf (txtIndex = CONFIG_CHECK_FAILURE) Then
                maxNumber = MAX_CHECK_FAILURE
              End If
            
              If (t.text >= minNumber And t.text <= maxNumber) Then
                hasError = False
              End If
            End If
          
            If (hasError) Then
              t.BackColor = TXT_ERROR_COLOR
              errors = errors + 1
              Exit For
            End If
          End If
        Next i
        
        Dim hexConfig() As String
          
        hexConfig = Split(CONFIG_HEX_INDEXES)
        
        For i = 0 To UBound(hexConfig)
          
          txtIndex = hexConfig(i)
        
          If (t.Index = txtIndex) Then
            hasError = True
          
            If (IsNumeric("&H" & t.text)) Then
              If (("&H" & t.text) > 0 And ("&H" & t.text) <= MAX_VERBYTE) Then
                hasError = False
              End If
            End If
            
            If (hasError) Then
              t.BackColor = TXT_ERROR_COLOR
              errors = errors + 1
              Exit For
            End If
          End If
        Next i
      End If
    Else
      If (TypeOf o Is ComboBox) Then
        Dim c As ComboBox
        Set c = o
        
        If (Len(c.text) < 3 Or Not isValidServerAddress(c.text)) Then
          c.BackColor = TXT_ERROR_COLOR
          errors = errors + 1
        End If
      End If
    End If
  Next
  
  If (Len(txtConfig(CONFIG_CDKEY_PROFILE).text) > 0 And Not _
    isValidCDKeyProfile(txtConfig(CONFIG_CDKEY_PROFILE).text)) Then
    txtConfig(CONFIG_CDKEY_PROFILE).BackColor = TXT_ERROR_COLOR
    errors = errors + 1
  End If
  
  markFormErrors = errors
End Function

Public Sub markErrorLocations(ByVal errors As Dictionary)
  Dim errorLocation As Variant
  
  For Each errorLocation In errors.Keys
    Dim str As String, txtControlIdx As Integer, isFill As Boolean, value As String
    isFill = False
    
    str = errorLocation
    value = errors.Item(errorLocation)
    
    If (Right(str, 1) = "f") Then
      txtControlIdx = left(str, Len(str) - 1)
      isFill = True
    Else
      txtControlIdx = str
    End If
    
    If (txtControlIdx = CONFIG_SERVER) Then
      cmbServer.text = value
      cmbServer.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_ERROR_COLOR)
    ElseIf (txtControlIdx = CONFIG_ADD_DATE_TO_TESTED) Then
      chkAddDateToTested.value = IIf(value, 1, 0)
      chkAddDateToTested.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
    ElseIf (txtControlIdx = CONFIG_SAVE_GOOD_PROXIES) Then
      chkSaveGoodProxies.value = IIf(value, 1, 0)
      chkSaveGoodProxies.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
    ElseIf (txtControlIdx = CONFIG_SKIP_FAILED_PROXIES) Then
      chkSkipFailedProxies.value = IIf(value, 1, 0)
      chkSkipFailedProxies.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
    ElseIf (txtControlIdx = CONFIG_ADD_REALM_TO_PROFILE) Then
      chkAddRealmToProfile.value = IIf(value, 1, 0)
      chkAddRealmToProfile.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
    ElseIf (txtControlIdx = CONFIG_SAVE_WINDOW_POSITION) Then
      chkSaveWindowPosition.value = IIf(value, 1, 0)
      chkSaveWindowPosition.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
    Else
      txtConfig(txtControlIdx).text = value
      txtConfig(txtControlIdx).BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_ERROR_COLOR)
    End If
  Next
End Sub
