VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2688
   ClientLeft      =   4284
   ClientTop       =   3228
   ClientWidth     =   3372
   ControlBox      =   0   'False
   HelpContextID   =   2
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2688
   ScaleWidth      =   3372
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      WhatsThisHelpID =   2
      Width           =   3132
   End
   Begin VB.TextBox txtLogin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   1
      Top             =   960
      WhatsThisHelpID =   2
      Width           =   3132
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      Caption         =   "What's Your Name?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   2
      Width           =   3132
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
    If txtLogin.Text = "" Then
        End
    Else
        frmMain.strName = txtLogin.Text
        frmLogin.Hide
    End If
End Sub

Private Sub Form_Activate()
    txtLogin.SetFocus
End Sub

