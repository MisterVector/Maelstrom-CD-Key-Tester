VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkCheckUpdateOnStartup 
      BackColor       =   &H00404040&
      Caption         =   "Check for Update on Startup"
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
      Left            =   5760
      TabIndex        =   41
      Top             =   3600
      Width           =   3255
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
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   5400
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
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   2520
      Width           =   2355
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   16
      Left            =   1920
      TabIndex        =   14
      Top             =   5040
      Width           =   3375
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
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   2880
      Width           =   2355
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   7
      Left            =   3840
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
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
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   2445
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   14
      Left            =   9120
      TabIndex        =   18
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   13
      Left            =   9120
      TabIndex        =   17
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   11
      Left            =   7440
      TabIndex        =   16
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   1440
      Width           =   2655
   End
   Begin VB.ComboBox cmbServer 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "frmConfig.frx":0000
      Left            =   2640
      List            =   "frmConfig.frx":0002
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   4
      Left            =   2640
      TabIndex        =   5
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   5
      Left            =   2640
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   9
      Left            =   8760
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   10
      Left            =   8760
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   8
      Left            =   8760
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtConfig 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   6
      Left            =   4560
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
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
      Left            =   5760
      TabIndex        =   9
      Top             =   3240
      Width           =   2655
   End
   Begin VB.PictureBox pbQuit 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9420
      Picture         =   "frmConfig.frx":0004
      ScaleHeight     =   450
      ScaleWidth      =   465
      TabIndex        =   23
      Top             =   -15
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
      Left            =   240
      TabIndex        =   40
      Top             =   5040
      Width           =   1695
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
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   39
      Top             =   4680
      Width           =   3615
   End
   Begin VB.Label lblConfig 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Key Test Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   19
      Left            =   360
      TabIndex        =   38
      Top             =   3960
      Width           =   4935
   End
   Begin VB.Label lblConfig 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Version Byte Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   37
      Top             =   3960
      Width           =   3855
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Warcraft II Version Byte:"
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
      Left            =   5760
      TabIndex        =   36
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Diablo II Version Byte:"
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
      Left            =   5760
      TabIndex        =   35
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "BNLS Server:"
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
      Left            =   5760
      TabIndex        =   34
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lblConfig 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "General Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      TabIndex        =   33
      Top             =   840
      Width           =   9375
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
      Left            =   1440
      TabIndex        =   32
      Top             =   1800
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
      Index           =   6
      Left            =   1440
      TabIndex        =   31
      Top             =   1440
      Width           =   1215
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
      Left            =   1560
      TabIndex        =   30
      Top             =   2160
      Width           =   855
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
      Left            =   600
      TabIndex        =   29
      Top             =   2880
      Width           =   1815
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
      Left            =   1440
      TabIndex        =   28
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Proxy Timeout Time:"
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
      Left            =   5760
      TabIndex        =   27
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Proxy Reconnect Time:"
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
      Left            =   5760
      TabIndex        =   26
      Top             =   2160
      Width           =   2535
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
      Left            =   5760
      TabIndex        =   25
      Top             =   1440
      Width           =   2295
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
      Left            =   3480
      TabIndex        =   24
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblRestoreDefaults 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Restore Defaults"
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
      Left            =   2280
      TabIndex        =   20
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
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
      Left            =   120
      TabIndex        =   19
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   7920
      TabIndex        =   21
      Top             =   6240
      Width           =   1815
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
      TabIndex        =   22
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const NUMERIC_TEXTBOX_IDENTIFIER As String = "numeric"
Private Const HEX_TEXTBOX_IDENTIFIER As String = "hex"
Private Const CONFIG_DEFAULT_TEXTBOX_IDXS = "5 6 7 8 9 10 11 13 14 15"

Private Sub chkAddDateToTested_Click()
    If (chkAddDateToTested.BackColor <> FRM_BACK_COLOR) Then
        chkAddDateToTested.BackColor = FRM_BACK_COLOR
    End If
End Sub

Private Sub chkAddDateToTested_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(Me, keyCode, Shift)
End Sub

Private Sub chkAddDateToTested_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub chkAddRealmToProfile_Click()
    If (chkAddRealmToProfile.BackColor <> FRM_BACK_COLOR) Then
        chkAddRealmToProfile.BackColor = FRM_BACK_COLOR
    End If
End Sub

Private Sub chkAddRealmToProfile_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(Me, keyCode, Shift)
End Sub

Private Sub chkAddRealmToProfile_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub chkCheckUpdateOnStartup_Click()
    If (chkCheckUpdateOnStartup.BackColor <> FRM_BACK_COLOR) Then
        chkCheckUpdateOnStartup.BackColor = FRM_BACK_COLOR
    End If
End Sub

Private Sub chkCheckUpdateOnStartup__KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(Me, keyCode, Shift)
End Sub

Private Sub chkCheckUpdateOnStartup_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub chkSaveGoodProxies_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(Me, keyCode, Shift)
End Sub

Private Sub chkSaveGoodProxies_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub chkSaveGoodProxies_Click()
    If (chkSaveGoodProxies.BackColor <> FRM_BACK_COLOR) Then
        chkSaveGoodProxies.BackColor = FRM_BACK_COLOR
    End If
End Sub

Private Sub chkSkipFailedProxies_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(Me, keyCode, Shift)
End Sub

Private Sub chkSkipFailedProxies_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub chkSkipFailedProxies_Click()
    If (chkSkipFailedProxies.BackColor <> FRM_BACK_COLOR) Then
        chkSkipFailedProxies.BackColor = FRM_BACK_COLOR
    End If
  
    chkSaveGoodProxies.Enabled = IIf(chkSkipFailedProxies.Value = 1, True, False)
End Sub

Private Sub chkSaveWindowPosition_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(Me, keyCode, Shift)
End Sub

Private Sub chkSaveWindowPosition_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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

Private Sub cmbServer_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(Me, keyCode, Shift)
End Sub

Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(Me, keyCode, Shift)
End Sub

Private Sub Form_Load()
    Dim gateway As Variant

    txtConfig(CONFIG_SOCKETS).Tag = NUMERIC_TEXTBOX_IDENTIFIER
    txtConfig(CONFIG_SOCKETS_PER_PROXY).Tag = NUMERIC_TEXTBOX_IDENTIFIER
    txtConfig(CONFIG_EXP_TESTS_PER_REG_KEY).Tag = NUMERIC_TEXTBOX_IDENTIFIER
    txtConfig(CONFIG_TEST_COUNT_PER_PROXY).Tag = NUMERIC_TEXTBOX_IDENTIFIER
    txtConfig(CONFIG_CHECK_FAILURE).Tag = NUMERIC_TEXTBOX_IDENTIFIER
    txtConfig(CONFIG_RECONNECT_TIME).Tag = NUMERIC_TEXTBOX_IDENTIFIER
    
    txtConfig(CONFIG_VERBYTE_W2BN).Tag = HEX_TEXTBOX_IDENTIFIER
    txtConfig(CONFIG_VERBYTE_D2DV).Tag = HEX_TEXTBOX_IDENTIFIER

    txtConfig(CONFIG_USERNAME).text = config.name
    txtConfig(CONFIG_PASSWORD).text = config.password
  
    If (config.server <> vbNullString) Then
        cmbServer.text = config.server
        cmbServer.AddItem config.server
    End If
  
    For Each gateway In dicGatewayIPs.keys
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
    
    chkAddDateToTested.Value = IIf(config.addDateToTested, 1, 0)
    chkSaveGoodProxies.Value = IIf(config.saveGoodProxies, 1, 0)
    chkSkipFailedProxies.Value = IIf(config.skipFailedProxies, 1, 0)
    chkAddRealmToProfile.Value = IIf(config.addRealmToProfile, 1, 0)
    chkSaveWindowPosition.Value = IIf(config.saveWindowPosition, 1, 0)
    chkCheckUpdateOnStartup.Value = IIf(config.checkUpdateOnStartup, 1, 0)
  
    If (chkSkipFailedProxies.Value = 0) Then
        chkSaveGoodProxies.Enabled = False
    End If
  
    txtConfig(CONFIG_BNLS_SERVER).text = config.bnlsServer
    
    txtConfig(CONFIG_VERBYTE_W2BN).text = Hex$(config.W2BNVerByte)
    txtConfig(CONFIG_VERBYTE_D2DV).text = Hex$(config.D2DVVerByte)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (isClosing) Then Exit Sub

    If (Not hasConfig) Then
        Dim msgBoxResult As Integer
        msgBoxResult = MsgBox("Are you sure you want to cancel without a valid configuration?", vbYesNo Or vbQuestion, PROGRAM_TITLE)
  
        If (msgBoxResult = vbNo) Then Cancel = 1
    End If
  
    frmMain.lblStart.Enabled = True
End Sub

Private Sub lblCancel_Click()
    Unload Me
End Sub

Private Sub lblCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub lblConfig_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub lblOk_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
    
    config.name = txtConfig(CONFIG_USERNAME).text
    config.password = txtConfig(CONFIG_PASSWORD).text
    
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
    
    config.addDateToTested = IIf(chkAddDateToTested.Value, True, False)
    config.saveGoodProxies = IIf(chkSaveGoodProxies.Value, True, False)
    config.skipFailedProxies = IIf(chkSkipFailedProxies.Value, True, False)
    config.addRealmToProfile = IIf(chkAddRealmToProfile.Value, True, False)
    config.saveWindowPosition = IIf(chkSaveWindowPosition.Value, True, False)
    config.checkUpdateOnStartup = IIf(chkCheckUpdateOnStartup.Value, True, False)
    
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
  
        frmMain.lblControl(CDKEY_PROFILE).Caption = "Using CD-Key Profile: " & fullProfileName
    Else
        frmMain.lblControl(CDKEY_PROFILE).Caption = "CD-Key Profile Not Configured"
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

    txtConfig(CONFIG_VERBYTE_W2BN).text = Hex$(DEFAULT_VERBYTE_W2BN)
    txtConfig(CONFIG_VERBYTE_D2DV).text = Hex$(DEFAULT_VERBYTE_D2DV)
    
    Dim defaultIndexes() As String
    defaultIndexes = Split(CONFIG_DEFAULT_TEXTBOX_IDXS, " ")
  
    For i = 0 To UBound(defaultIndexes)
        txtConfig(defaultIndexes(i)).BackColor = TXT_BACK_COLOR
    Next i
  
    chkAddDateToTested.Value = IIf(DEFAULT_ADD_DATE_TO_TESTED, 1, 0)
    chkSaveGoodProxies.Value = IIf(DEFAULT_SAVE_GOOD_PROXIES, 1, 0)
    chkSkipFailedProxies.Value = IIf(DEFAULT_SKIP_FAILED_PROXIES, 1, 0)
    chkAddRealmToProfile.Value = IIf(DEFAULT_ADD_REALM_TO_PROFILE, 1, 0)
    chkSaveWindowPosition.Value = IIf(DEFAULT_SAVE_WINDOW_POSITION, 1, 0)
    chkCheckUpdateOnStartup.Value = IIf(DEFAULT_CHECK_UPDATE_ON_STARTUP, 1, 0)
    
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
    
    If (chkCheckUpdateOnStartup.BackColor <> FRM_BACK_COLOR) Then
        chkCheckUpdateOnStartup.BackColor = FRM_BACK_COLOR
    End If
    
    chkSaveGoodProxies.Enabled = IIf(chkSkipFailedProxies.Value = 1, True, False)
End Sub

Private Sub lblRestoreDefaults_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call moveEntireForm(Me, Button)
End Sub

Private Sub pbQuit_KeyDown(keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(Me, keyCode, Shift)
End Sub

Private Sub pbQuit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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

Private Sub txtConfig_KeyDown(Index As Integer, keyCode As Integer, Shift As Integer)
    Call checkForQuitShortcut(Me, keyCode, Shift)
End Sub

Private Function markFormErrors() As Integer
    Dim o As Control, errors As Integer
  
    For Each o In Me.Controls
        If (TypeOf o Is TextBox) Then
            Dim t As TextBox, lenRequire As Integer, hasError As Boolean
            hasError = False
      
            Set t = o
            lenRequire = IIf(t.Index = CONFIG_USERNAME Or t.Index = CONFIG_BNLS_SERVER, 3, IIf(t.Index = CONFIG_CDKEY_PROFILE, 0, 1))

            If (Len(t.text) < lenRequire) Then
                t.BackColor = TXT_ERROR_COLOR
                errors = errors + 1
                hasError = True
            End If

            If (Not hasError) Then
                If (t.Tag = NUMERIC_TEXTBOX_IDENTIFIER) Then
                    If (IsNumericB(t.text)) Then
                        Dim minNumber As Integer, maxNumber As Integer

                        If (t.Index = CONFIG_TEST_COUNT_PER_PROXY) Then
                            minNumber = 0
                        Else
                            minNumber = 1
                        End If

                        If (t.Index = CONFIG_SOCKETS) Then
                            maxNumber = MAX_SOCKETS
                        ElseIf (t.Index = CONFIG_SOCKETS_PER_PROXY) Then
                            maxNumber = MAX_SOCKETS_PER_PROXY
                        ElseIf (t.Index = CONFIG_EXP_TESTS_PER_REG_KEY) Then
                            maxNumber = MAX_EXP_TESTS_PER_REG_KEY
                        ElseIf (t.Index = CONFIG_TEST_COUNT_PER_PROXY) Then
                            maxNumber = MAX_TEST_COUNT_PER_PROXY
                        ElseIf (t.Index = CONFIG_RECONNECT_TIME) Then
                          maxNumber = MAX_RECONNECT_TIME
                        ElseIf (t.Index = CONFIG_CHECK_FAILURE) Then
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

                If (t.Tag = HEX_TEXTBOX_IDENTIFIER) Then
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
  
    For Each errorLocation In errors.keys
        Dim str As String, txtControlIdx As Integer, isFill As Boolean, Value As String
        isFill = False
    
        str = errorLocation
        Value = errors.Item(errorLocation)
    
        If (Right$(str, 1) = "f") Then
            txtControlIdx = left$(str, Len(str) - 1)
            isFill = True
        Else
            txtControlIdx = str
        End If
    
        If (txtControlIdx = CONFIG_SERVER) Then
            cmbServer.text = Value
            cmbServer.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_ERROR_COLOR)
        ElseIf (txtControlIdx = CONFIG_ADD_DATE_TO_TESTED) Then
            chkAddDateToTested.Value = IIf(Value, 1, 0)
            chkAddDateToTested.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
        ElseIf (txtControlIdx = CONFIG_SAVE_GOOD_PROXIES) Then
            chkSaveGoodProxies.Value = IIf(Value, 1, 0)
            chkSaveGoodProxies.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
        ElseIf (txtControlIdx = CONFIG_SKIP_FAILED_PROXIES) Then
            chkSkipFailedProxies.Value = IIf(Value, 1, 0)
            chkSkipFailedProxies.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
        ElseIf (txtControlIdx = CONFIG_ADD_REALM_TO_PROFILE) Then
            chkAddRealmToProfile.Value = IIf(Value, 1, 0)
            chkAddRealmToProfile.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
        ElseIf (txtControlIdx = CONFIG_SAVE_WINDOW_POSITION) Then
            chkSaveWindowPosition.Value = IIf(Value, 1, 0)
            chkSaveWindowPosition.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
        ElseIf (txtControlIdx = CONFIG_CHECK_UPDATE_ON_STARTUP) Then
            chkCheckUpdateOnStartup.Value = IIf(Value, 1, 0)
            chkCheckUpdateOnStartup.BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_WARN_COLOR)
        Else
            txtConfig(txtControlIdx).text = Value
            txtConfig(txtControlIdx).BackColor = IIf(isFill, TXT_FILL_COLOR, TXT_ERROR_COLOR)
        End If
    Next
End Sub
