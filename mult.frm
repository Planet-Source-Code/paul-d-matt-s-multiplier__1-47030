VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5652
   ClientLeft      =   4128
   ClientTop       =   1584
   ClientWidth     =   3984
   ForeColor       =   &H00000000&
   HelpContextID   =   1
   Icon            =   "mult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5652
   ScaleWidth      =   3984
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   3480
      Top             =   5160
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      HelpFile        =   "Matt.hlp"
   End
   Begin VB.Frame fraScores 
      Caption         =   "Last Timed Score"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   3840
      WhatsThisHelpID =   10
      Width           =   3732
      Begin VB.Label lblHSRightLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Right"
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
         Index           =   1
         Left            =   228
         TabIndex        =   33
         Top             =   840
         WhatsThisHelpID =   10
         Width           =   516
      End
      Begin VB.Label lblHSWrongLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Wrong"
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
         Index           =   1
         Left            =   1056
         TabIndex        =   32
         Top             =   840
         WhatsThisHelpID =   10
         Width           =   660
      End
      Begin VB.Label lblHSRight 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   384
         Index           =   1
         Left            =   324
         TabIndex        =   31
         Top             =   360
         WhatsThisHelpID =   10
         Width           =   192
      End
      Begin VB.Label lblHSWrong 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   384
         Index           =   1
         Left            =   1272
         TabIndex        =   30
         Top             =   360
         WhatsThisHelpID =   10
         Width           =   204
      End
      Begin VB.Label lblHSPercent 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   384
         Index           =   1
         Left            =   2244
         TabIndex        =   29
         Top             =   360
         WhatsThisHelpID =   10
         Width           =   204
      End
      Begin VB.Label lblHSPercentLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Percent"
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
         Index           =   1
         Left            =   2028
         TabIndex        =   28
         Top             =   840
         WhatsThisHelpID =   10
         Width           =   756
      End
      Begin VB.Label lblHSTotal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Index           =   1
         Left            =   3204
         TabIndex        =   27
         Top             =   360
         WhatsThisHelpID =   10
         Width           =   204
      End
      Begin VB.Label lblHSTotallabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Index           =   1
         Left            =   3060
         TabIndex        =   26
         Top             =   840
         WhatsThisHelpID =   10
         Width           =   492
      End
      Begin VB.Label lblHSDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   252
         Index           =   1
         Left            =   1080
         TabIndex        =   25
         Top             =   1320
         WhatsThisHelpID =   10
         Width           =   2532
      End
      Begin VB.Label lblWhenLabel 
         Caption         =   "When:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   1320
         WhatsThisHelpID =   10
         Width           =   612
      End
   End
   Begin VB.Timer tmrTimedTest 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   3600
   End
   Begin VB.CommandButton cmdTimed 
      Height          =   732
      HelpContextID   =   5
      Left            =   1440
      Picture         =   "mult.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      WhatsThisHelpID =   5
      Width           =   1092
   End
   Begin VB.Frame fraScores 
      Caption         =   "High Timed Score"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   3732
      Begin VB.Label lblWhenLabel 
         Caption         =   "When:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         WhatsThisHelpID =   10
         Width           =   612
      End
      Begin VB.Label lblHSDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   252
         Index           =   0
         Left            =   1080
         TabIndex        =   21
         Top             =   1320
         WhatsThisHelpID =   10
         Width           =   2532
      End
      Begin VB.Label lblHSTotallabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Index           =   0
         Left            =   3060
         TabIndex        =   16
         Top             =   840
         WhatsThisHelpID =   10
         Width           =   492
      End
      Begin VB.Label lblHSTotal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Index           =   0
         Left            =   3204
         TabIndex        =   15
         Top             =   360
         WhatsThisHelpID =   10
         Width           =   204
      End
      Begin VB.Label lblHSPercentLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Percent"
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
         Index           =   0
         Left            =   2028
         TabIndex        =   13
         Top             =   840
         WhatsThisHelpID =   10
         Width           =   756
      End
      Begin VB.Label lblHSPercent 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   384
         Index           =   0
         Left            =   2244
         TabIndex        =   12
         Top             =   360
         WhatsThisHelpID =   10
         Width           =   204
      End
      Begin VB.Label lblHSWrong 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   384
         Index           =   0
         Left            =   1272
         TabIndex        =   11
         Top             =   360
         WhatsThisHelpID =   10
         Width           =   204
      End
      Begin VB.Label lblHSRight 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   384
         Index           =   0
         Left            =   324
         TabIndex        =   10
         Top             =   360
         WhatsThisHelpID =   10
         Width           =   192
      End
      Begin VB.Label lblHSWrongLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Wrong"
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
         Index           =   0
         Left            =   1056
         TabIndex        =   9
         Top             =   840
         WhatsThisHelpID =   10
         Width           =   660
      End
      Begin VB.Label lblHSRightLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Right"
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
         Index           =   0
         Left            =   228
         TabIndex        =   8
         Top             =   840
         WhatsThisHelpID =   10
         Width           =   516
      End
   End
   Begin VB.CommandButton cmdEnter 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   1440
      MaskColor       =   &H80000004&
      Picture         =   "mult.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      WhatsThisHelpID =   10
      Width           =   1092
   End
   Begin VB.TextBox txtAnswer 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   504
      Left            =   1440
      TabIndex        =   0
      Top             =   1908
      WhatsThisHelpID =   10
      Width           =   1092
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   552
      Left            =   3168
      TabIndex        =   20
      Top             =   3120
      WhatsThisHelpID =   10
      Width           =   252
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   552
      Left            =   528
      TabIndex        =   19
      Top             =   3120
      WhatsThisHelpID =   10
      Width           =   252
   End
   Begin VB.Label lblTotalLabel 
      Alignment       =   2  'Center
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2640
      TabIndex        =   18
      Top             =   2520
      WhatsThisHelpID =   10
      Width           =   1212
   End
   Begin VB.Label lblPercentLabel 
      Alignment       =   2  'Center
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      WhatsThisHelpID =   10
      Width           =   1092
   End
   Begin VB.Label lblWrongLabel 
      Alignment       =   2  'Center
      Caption         =   "Wrong"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2640
      TabIndex        =   6
      Top             =   1200
      WhatsThisHelpID =   10
      Width           =   1212
   End
   Begin VB.Label lblRightLabel 
      Alignment       =   2  'Center
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      WhatsThisHelpID =   10
      Width           =   1212
   End
   Begin VB.Label lblWrong 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   552
      Left            =   3180
      TabIndex        =   4
      Top             =   1824
      WhatsThisHelpID =   10
      Width           =   252
   End
   Begin VB.Label lblRight 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   552
      Left            =   540
      TabIndex        =   3
      Top             =   1800
      WhatsThisHelpID =   10
      Width           =   252
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   28.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   804
      Left            =   156
      TabIndex        =   1
      Top             =   120
      WhatsThisHelpID =   10
      Width           =   3660
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   210
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         HelpContextID   =   3
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      HelpContextID   =   220
      Begin VB.Menu mnuDing 
         Caption         =   "&Timer Ding"
         Checked         =   -1  'True
         HelpContextID   =   3
      End
      Begin VB.Menu mnuShowTips 
         Caption         =   "&Show Tips"
         Checked         =   -1  'True
         HelpContextID   =   3
      End
      Begin VB.Menu mnuRemoveOnes 
         Caption         =   "R&emove 1's"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset"
         HelpContextID   =   3
      End
      Begin VB.Menu mnuPickTables 
         Caption         =   "&Pick Table..."
         HelpContextID   =   3
      End
   End
   Begin VB.Menu mnuReview 
      Caption         =   "Re&view"
      HelpContextID   =   230
      Begin VB.Menu mnuTry 
         Caption         =   "&Try Again"
         Checked         =   -1  'True
         HelpContextID   =   3
      End
      Begin VB.Menu mnuClearReview 
         Caption         =   "Clear"
         HelpContextID   =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents && Index"
         HelpContextID   =   1
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Matt 's Multiplier
'June 2003
'© Paul Dahlen, 2003

Const intSeconds = 60                                 'set timer seconds down during debug
Const HelpCNT = &HB
Dim intAnswer As Integer
'Number Right & Wrong
Dim intRight As Integer
Dim intWrong As Integer
Dim intPercent As Integer
Dim intTotal As Integer
'Number Right & Wrong from Last Timed Test
Dim intSavedRight As Integer
Dim intSavedWrong As Integer
Dim intSavedPercent As Integer
Dim intSavedTotal As Integer
'Number Right & Wrong from High Score Timed Test
Dim intHSSavedRight As Integer
Dim intHSSavedWrong As Integer
Dim intHSSavedPercent As Integer
Dim intHSSavedTotal As Integer
'Counter counts down seconds & mult2counter counts up through a times table
Dim intCounter As Integer
Dim intMult2Counter As Integer
'Review arrays
Dim intReview1() As Integer
Dim intReview2() As Integer
Dim intReviewCounter As Integer
Dim intReviewIndex As Integer
'Switches for Reviewing, Showing Answer & Timer Dinging
Dim blnReview As Boolean
Dim blnShowAnswer As Boolean
Dim blnDing As Boolean
Dim blnRemoveOnes As Boolean
'Strings for registry settings
Dim strDing As String
Dim strShowHints As String
Dim strRemoveOnes As String
Public strShowAbout As String
Public strName As String
'Multiplier and multiplicand
Public intMult1 As Integer
Public intMult2 As Integer
'Switch for Random multiplier and multiplicand
Public blnRandomNumbers As Boolean
'initialization
Private Sub Form_Load()
'Get login Name, put name in Public string variable
    frmLogin.Show vbModal, frmMain
    frmMain.Caption = strName
    'get registry settings
    GetRegistry
    'set options from registry settings
    SetOptions
    'High Score timed test caption values to High Score timed test integer values
    intHSSavedRight = Int(lblHSRight(0).Caption)
    intHSSavedWrong = Int(lblHSWrong(0).Caption)
    intHSSavedPercent = Int(lblHSPercent(0).Caption)
    intHSSavedTotal = Int(lblHSTotal(0).Caption)
    'last timed test caption values to last timed test integer values
    intSavedRight = Int(lblHSRight(1).Caption)
    intSavedWrong = Int(lblHSWrong(1).Caption)
    intSavedPercent = Int(lblHSPercent(1).Caption)
    intSavedTotal = Int(lblHSTotal(1).Caption)
    'set random numbers on
    blnRandomNumbers = True
    'set review off
    blnReview = False
    'set default to show answers when wrong
    blnShowAnswer = True
    'set review counter to zero & turn off check in menu
    mnuTry.Checked = False
    mnuTry.Enabled = False
    intReviewCounter = 0
    intReviewIndex = 0
    ReDim intReview1(intReviewCounter + 1)
    ReDim intReview2(intReviewCounter + 1)
    ' Set Score to Zero
    Clear
    'show first multiplication problem to solve
    ShowProblem
    'show About form unless checked
    If strShowAbout = "Show" Then frmAbout.Show vbModal, frmMain
    
End Sub
'save settings and end
Private Sub Form_Unload(Cancel As Integer)
'Make sure % doesn't try to divide by zero
    If intRight + intWrong > 0 Then
        intPercent = (intRight / (intRight + intWrong)) * 100
    Else
        intPercent = 0
    End If
    SaveRegistry
End Sub
'Enter Button
Private Sub cmdEnter_Click()
'if Answer Box is blank, prompt
    If txtAnswer = "" Then
        txtAnswer = "0"
        MsgBox "Put a number in!"
    End If
    'Right or Wrong
    If txtAnswer.Text = intAnswer Then
        intRight = intRight + 1
        lblRight.Caption = intRight
    Else:
        intWrong = intWrong + 1
        lblWrong.Caption = Str(intWrong)
        'show answer if not in timed round
        If blnShowAnswer Then message = MsgBox(intMult1 & " X " & intMult2 & " = " & intAnswer, , "Whoops!")
        mnuTry.Enabled = True
        If Not mnuTry.Checked Then
            intReview1(intReviewCounter) = intMult1
            intReview2(intReviewCounter) = intMult2
            ReDim Preserve intReview2(UBound(intReview2) + 1)
            ReDim Preserve intReview1(UBound(intReview1) + 1)
            intReviewCounter = intReviewCounter + 1
        End If
    End If
    'set Total and % Numbers on Form
    lblTotal.Caption = intRight + intWrong
    lblPercent.Caption = Int((intRight / (intRight + intWrong)) * 100)
    'show next multiplication problem to solve
    ShowProblem
    'set focus to textbox and clear
    txtAnswer.Text = ""
    txtAnswer.SetFocus
End Sub
'Timer Button
Private Sub cmdTimed_Click()
'set timer from const intSeconds value at top
    intCounter = intSeconds
    'gray out timer button
    cmdTimed.Enabled = False
    'turn on timer function
    tmrTimedTest.Enabled = True
    Clear
    'do not show correct answers during test
    blnShowAnswer = False
    blnRandomNumbers = True
    ShowProblem
    txtAnswer.SetFocus
End Sub
'View Last Score or High Score with click on Frame 0-High, 1-Last
Private Sub fraScores_Click(index As Integer)
    Select Case index
    Case 0
        fraScores(0).Visible = False
        fraScores(1).Visible = True
    Case 1
        fraScores(1).Visible = False
        fraScores(0).Visible = True
    End Select
End Sub
'turn timer dinger on or off
Private Sub mnuDing_Click()
    If Not tmrTimedTest.Enabled Then
        If mnuDing.Checked Then
            mnuDing.Checked = False
            blnDing = False
            strDing = "off"
        Else
            mnuDing.Checked = True
            blnDing = True
            strDing = "on"
        End If
    End If
End Sub
'exit program
Private Sub mnuExit_Click()
    Unload Me
    End
End Sub
'show About form
Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub
'Show Helpfile Index
Private Sub mnuHelpContents_Click()
    With dlgCommon
        ' You must set the Help file name.
        .HelpFile = "Matt.hlp"
        ' Display the Table of Contents. Note that the
        ' HelpCNT contstant is not an intrinsic
        ' constant. The cdlHelpSetContents ensures that
        ' only the Table of Contents (not Index or Find)
        ' shows.
        .HelpCommand = HelpCNT Or cdlHelpSetContents
        .ShowHelp
    End With
End Sub
'pick tables
Private Sub mnuPickTables_Click()
    frmTables.Show vbModal, frmMain
    intMult2Counter = 2
    ShowProblem
End Sub
'do not include the number 1 in the random numbers
Private Sub mnuRemoveOnes_Click()
    If mnuRemoveOnes.Checked Then
        mnuRemoveOnes.Checked = False
        blnRemoveOnes = False
    Else
        mnuRemoveOnes.Checked = True
        blnRemoveOnes = True
        strRemoveOnes = "on"
    End If
    Clear
    ShowProblem
End Sub
'reset
Private Sub mnuReset_Click()
    cmdEnter.Enabled = True
    cmdTimed.Enabled = True
    txtAnswer.Enabled = True
    txtAnswer.Text = ""
    txtAnswer.SetFocus
    blnShowAnswer = True
    Clear
    ShowProblem
End Sub
'show tips
Private Sub mnuShowTips_Click()
    If mnuShowTips.Checked Then
        mnuShowTips.Checked = False
        cmdEnter.ToolTipText = ""
        cmdTimed.ToolTipText = ""
        fraScores(0).ToolTipText = ""
        fraScores(1).ToolTipText = ""
        strShowHints = "off"
    Else
        mnuShowTips.Checked = True
        cmdEnter.ToolTipText = "Click to enter your answer"
        cmdTimed.ToolTipText = "Click to start a timed test"
        fraScores(0).ToolTipText = "Click to see last score"
        fraScores(1).ToolTipText = "Click to see high score"
        strShowHints = "on"
    End If

End Sub
'try problems that you got wrong
Private Sub mnuTry_Click()
    If mnuTry.Checked Then
        mnuTry.Checked = False
        blnReview = False
    Else
        mnuTry.Checked = True
        blnReview = True
    End If
    Clear
    ShowProblem
End Sub
'clear review integer arrays
Private Sub mnuClearReview_Click()
    mnuTry.Checked = False
    mnuTry.Enabled = False
    blnReview = False
    intReviewCounter = 0
    intReviewIndex = 0
    ReDim intReview1(intReviewCounter + 1)
    ReDim intReview2(intReviewCounter + 1)
    Clear
    ShowProblem
End Sub
'Timer subroutine
Private Sub tmrTimedTest_Timer()
'count down 1 minute by seconds in the taskbar - ding either on or off
    frmMain.Caption = intCounter
    intCounter = intCounter - 1
    If blnDing Then Beep
    If intCounter = 0 Then
        cmdEnter.Enabled = False
        tmrTimedTest.Enabled = False
        frmMain.Caption = strName
        tmrTimedTest.Enabled = False
        txtAnswer.Enabled = False
        Score
    End If
End Sub
'pick 2 numbers-either random or specific tables
Sub ShowProblem()
    If blnReview Then
        If intReviewIndex = intReviewCounter Then intReviewIndex = 0
        intMult1 = intReview1(intReviewIndex)
        intMult2 = intReview2(intReviewIndex)
        intReviewIndex = intReviewIndex + 1
    ElseIf blnRandomNumbers Then
        Randomize
        ' Generate random values between 2 and 12 or 1 and 12.
        If blnRemoveOnes Then intMult1 = Int((11 * Rnd) + 2) Else intMult1 = Int((12 * Rnd) + 1)
        Randomize
        If blnRemoveOnes Then intMult2 = Int((11 * Rnd) + 2) Else intMult2 = Int((12 * Rnd) + 1)
    Else
        'intMult1 was set in the Tables Form
        If intMult2Counter = 13 Then intMult2Counter = 2
        intMult2 = intMult2Counter
        intMult2Counter = intMult2Counter + 1
    End If
    intAnswer = intMult1 * intMult2
    lblQuestion.Caption = Str(intMult1) & "  X " & Str(intMult2) & " ="
End Sub
'Clear scores
Sub Clear()
    intRight = 0
    intWrong = 0
    intPercent = 0
    itotal = 0
    intMult2Counter = 2
    lblRight.Caption = "0"
    lblWrong.Caption = "0"
    lblPercent.Caption = "0"
    lblTotal.Caption = "0"
End Sub
'Check to see if High Score if so - post
Public Sub Score()
'Make sure % doesn't try to divide by zero
    If intRight + intWrong > 0 Then
        intPercent = Int((intRight / (intRight + intWrong)) * 100)
    Else
        intPercent = 0
    End If
    'if total right is higher and the percent is 90 or above - new top score
    'Index 0 is High Score - Index 1 is Last Score
    If intRight >= intHSSavedRight And intPercent >= 90 Then
        intHSSavedRight = intRight
        lblHSRight(0).Caption = intRight
        intHSSavedWrong = intWrong
        lblHSWrong(0).Caption = intWrong
        intHSSavedPercent = intPercent
        lblHSPercent(0).Caption = intPercent
        intHSSavedTotal = intHSSavedRight + intHSSavedWrong
        lblHSTotal(0).Caption = intHSSavedTotal
        lblHSDate(0).Caption = Format(Now, "dddd, mmmm dd")
    End If
    'always save last score
    intSavedRight = intRight
    lblHSRight(1).Caption = intRight
    intSavedWrong = intWrong
    frmMain.lblHSWrong(1).Caption = intWrong
    intSavedPercent = intPercent
    frmMain.lblHSPercent(1).Caption = intPercent
    intSavedTotal = intRight + intWrong
    lblHSTotal(1).Caption = intSavedTotal
    lblHSDate(1).Caption = Format(Now, "dddd, mmmm dd")

End Sub
Sub GetRegistry()
'get high score settings from registry and load them into HS integers - 0 if key not there
'Index 0 is High Score
    lblHSRight(0).Caption = GetSetting("Matt's Multipliers", strName, "HSRight", "0")
    lblHSWrong(0).Caption = GetSetting("Matt's Multipliers", strName, "HSWrong", "0")
    lblHSPercent(0).Caption = GetSetting("Matt's Multipliers", strName, "HSPercent", "0")
    lblHSTotal(0).Caption = GetSetting("Matt's Multipliers", strName, "HSTotal", "0")
    lblHSDate(0).Caption = GetSetting("Matt's Multipliers", strName, "HSDate", "")
    'Index 1 is Last Score
    lblHSRight(1).Caption = GetSetting("Matt's Multipliers", strName, "Right", "0")
    lblHSWrong(1).Caption = GetSetting("Matt's Multipliers", strName, "Wrong", "0")
    lblHSPercent(1).Caption = GetSetting("Matt's Multipliers", strName, "Percent", "0")
    lblHSTotal(1).Caption = GetSetting("Matt's Multipliers", strName, "Total", "0")
    lblHSDate(1).Caption = GetSetting("Matt's Multipliers", strName, "Date", "")
    'Get Program Settings
    strDing = GetSetting("Matt's Multipliers", strName & "\ Settings", "Dinger", "off")
    strShowHints = GetSetting("Matt's Multipliers", strName & "\ Settings", "Hints", "on")
    strRemoveOnes = GetSetting("Matt's Multipliers", strName & "\ Settings", "Ones", "on")
    strShowAbout = GetSetting("Matt's Multipliers", strName & "\ Settings", "About", "Show")
End Sub
Sub SaveRegistry()
'Save HighScore Settings HKCU/Software/VB & VBA Program Settings
'Index 0 is High Score
    SaveSetting "Matt's Multipliers", strName, "HSRight", intHSSavedRight
    SaveSetting "Matt's Multipliers", strName, "HSWrong", intHSSavedWrong
    SaveSetting "Matt's Multipliers", strName, "HSPercent", intHSSavedPercent
    SaveSetting "Matt's Multipliers", strName, "HSTotal", intHSSavedTotal
    SaveSetting "Matt's Multipliers", strName, "HSDate", lblHSDate(0).Caption
    'Index 1 is Last Score
    SaveSetting "Matt's Multipliers", strName, "Right", intSavedRight
    SaveSetting "Matt's Multipliers", strName, "Wrong", intSavedWrong
    SaveSetting "Matt's Multipliers", strName, "Percent", intSavedPercent
    SaveSetting "Matt's Multipliers", strName, "Total", intSavedTotal
    SaveSetting "Matt's Multipliers", strName, "Date", lblHSDate(1).Caption
    'Program Settings
    SaveSetting "Matt's Multipliers", strName & "\ Settings", "Dinger", strDing
    SaveSetting "Matt's Multipliers", strName & "\ Settings", "Hints", strShowHints
    SaveSetting "Matt's Multipliers", strName & "\ Settings", "Ones", strRemoveOnes
    SaveSetting "Matt's Multipliers", strName & "\ Settings", "About", strShowAbout
End Sub

Sub SetOptions()
'Set Timer Dinger
    If strDing = "on" Then
        mnuDing.Checked = True
        blnDing = True
    Else:
        mnuDing.Checked = False
        blnDing = False
    End If
    If strShowHints = "on" Then
        mnuShowTips.Checked = True
        cmdEnter.ToolTipText = "Click to enter your answer"
        cmdTimed.ToolTipText = "Click to start a timed test"
        fraScores(0).ToolTipText = "Click to see last score"
        fraScores(1).ToolTipText = "Click to see high score"
    Else:
        mnuShowTips.Checked = False
        cmdEnter.ToolTipText = ""
        cmdTimed.ToolTipText = ""
        fraScores(0).ToolTipText = ""
        fraScores(1).ToolTipText = ""
    End If
    If strRemoveOnes = "on" Then
        mnuRemoveOnes.Checked = True
        blnRemoveOnes = True
    Else
        mnuRemoveOnes.Checked = False
        blnRemoveOnes = False
    End If
End Sub
