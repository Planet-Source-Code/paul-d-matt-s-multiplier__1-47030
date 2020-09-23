VERSION 5.00
Begin VB.Form frmTables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Pick Table"
   ClientHeight    =   3600
   ClientLeft      =   6648
   ClientTop       =   4164
   ClientWidth     =   1824
   HelpContextID   =   240
   Icon            =   "Tables.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   1824
   Begin VB.OptionButton optTables 
      Caption         =   "Random"
      Height          =   372
      Index           =   11
      Left            =   960
      TabIndex        =   12
      Top             =   2640
      Value           =   -1  'True
      WhatsThisHelpID =   240
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   480
      TabIndex        =   11
      Top             =   3120
      WhatsThisHelpID =   240
      Width           =   852
   End
   Begin VB.OptionButton optTables 
      Caption         =   "12s"
      Height          =   372
      Index           =   10
      Left            =   960
      TabIndex        =   10
      Top             =   2160
      WhatsThisHelpID =   240
      Width           =   612
   End
   Begin VB.OptionButton optTables 
      Caption         =   "11s"
      Height          =   372
      Index           =   9
      Left            =   960
      TabIndex        =   9
      Top             =   1680
      WhatsThisHelpID =   240
      Width           =   612
   End
   Begin VB.OptionButton optTables 
      Caption         =   "10s"
      Height          =   372
      Index           =   8
      Left            =   960
      TabIndex        =   8
      Top             =   1200
      WhatsThisHelpID =   240
      Width           =   612
   End
   Begin VB.OptionButton optTables 
      Caption         =   "9s"
      Height          =   372
      Index           =   7
      Left            =   960
      TabIndex        =   7
      Top             =   720
      WhatsThisHelpID =   240
      Width           =   612
   End
   Begin VB.OptionButton optTables 
      Caption         =   "8s"
      Height          =   372
      Index           =   6
      Left            =   960
      TabIndex        =   6
      Top             =   240
      WhatsThisHelpID =   240
      Width           =   612
   End
   Begin VB.OptionButton optTables 
      Caption         =   "7s"
      Height          =   372
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      WhatsThisHelpID =   240
      Width           =   612
   End
   Begin VB.OptionButton optTables 
      Caption         =   "6s"
      Height          =   372
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      WhatsThisHelpID =   240
      Width           =   612
   End
   Begin VB.OptionButton optTables 
      Caption         =   "5s"
      Height          =   372
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      WhatsThisHelpID =   240
      Width           =   612
   End
   Begin VB.OptionButton optTables 
      Caption         =   "4s"
      Height          =   372
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      WhatsThisHelpID =   240
      Width           =   612
   End
   Begin VB.OptionButton optTables 
      Caption         =   "3s"
      Height          =   372
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   720
      WhatsThisHelpID =   240
      Width           =   612
   End
   Begin VB.OptionButton optTables 
      Caption         =   "2s"
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      WhatsThisHelpID =   240
      Width           =   612
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    For index = 0 To 10
        If optTables(index).Value Then
            frmMain.intMult1 = index + 2
            frmMain.blnRandomNumbers = False
        End If
    Next
    If optTables(11).Value Then frmMain.blnRandomNumbers = True
    frmTables.Hide
End Sub

Private Sub Form_Activate()
'if timed test was taken, random numbers option should be clicked on form
    If frmMain.blnRandomNumbers Then optTables(11).Value = True
End Sub
