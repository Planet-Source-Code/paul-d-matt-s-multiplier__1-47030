VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6108
   ClientLeft      =   2340
   ClientTop       =   1932
   ClientWidth     =   5736
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4205.204
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkDontShow 
      Caption         =   "Don't show this message on startup"
      Height          =   192
      Left            =   240
      TabIndex        =   6
      Top             =   5880
      Width           =   2892
   End
   Begin VB.PictureBox picDisclaimer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3732
      Left            =   240
      ScaleHeight     =   3732
      ScaleWidth      =   5292
      TabIndex        =   5
      Top             =   1920
      Width           =   5292
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   384
      Left            =   240
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   263.118
      ScaleMode       =   0  'User
      ScaleWidth      =   263.118
      TabIndex        =   1
      Top             =   240
      Width           =   384
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   1188
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   123.825
      X2              =   5210.038
      Y1              =   3982.138
      Y2              =   3982.138
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   112.568
      X2              =   5210.976
      Y1              =   3965.615
      Y2              =   3965.615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   90.055
      X2              =   5210.038
      Y1              =   1107.067
      Y2              =   1107.067
   End
   Begin VB.Label lblDescription 
      Caption         =   "© 2003, Paul R. Dahlen"
      ForeColor       =   &H00000000&
      Height          =   324
      Left            =   1056
      TabIndex        =   2
      Top             =   1248
      Width           =   3396
   End
   Begin VB.Label lblTitle 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   480
      Left            =   1056
      TabIndex        =   3
      Top             =   240
      Width           =   3276
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   101.312
      X2              =   5210.976
      Y1              =   1123.591
      Y2              =   1123.591
   End
   Begin VB.Label lblVersion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   1056
      TabIndex        =   4
      Top             =   780
      Width           =   3396
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
        KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
        KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                                      ' Unicode nul terminated string
Const REG_DWORD = 4                                   ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    picDisclaimer.Print "While I make every effort to deliver high quality products, I do not "
    picDisclaimer.Print "guarantee that my software free from defects.  My software is provided "
    picDisclaimer.Print " as is, and you use the software at your own risk."
    picDisclaimer.Print ""
    picDisclaimer.Print "I make no warranties as to performance, merchantability, fitness for"
    picDisclaimer.Print " a particular purpose, or any other warranties whether expressed"
    picDisclaimer.Print " or implied."
    picDisclaimer.Print ""
    picDisclaimer.Print "No oral or written communication from or informationprovided by me "
    picDisclaimer.Print "shall create a warranty.   "
    picDisclaimer.Print ""
    picDisclaimer.Print "Under no circumstances shall I be liable for direct, indirect, special, "
    picDisclaimer.Print "incidental, or consequential damages resulting from the use, misuse,"
    picDisclaimer.Print "or inability to use this software, even if I have been advised of the"
    picDisclaimer.Print "possibility of such damages. "
    picDisclaimer.Print ""
    picDisclaimer.Print "These exclusions and limitations may not apply in all jurisdictions. You "
    picDisclaimer.Print "may have additional rights and some of these limitations may not apply"
    picDisclaimer.Print "to you."
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title

End Sub


