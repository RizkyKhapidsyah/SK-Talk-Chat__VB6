VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic Speech Recognition"
   ClientHeight    =   6885
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9885
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   459
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRepeat 
      Caption         =   "Repeat Last"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   4920
      TabIndex        =   19
      Top             =   2640
      Width           =   4455
      Begin VB.ListBox lstConnected 
         Height          =   2205
         Left            =   2280
         TabIndex        =   26
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdKick 
         Caption         =   "Kick"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdJoin 
         Caption         =   "Join"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdHost 
         Caption         =   "Host"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frIcon 
      Caption         =   "Connection Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   4455
      Begin VB.TextBox txtHandle 
         Height          =   375
         Left            =   1800
         TabIndex        =   24
         Text            =   "JustAGuy"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtLocalIP 
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtLocalPort 
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtRemotePort 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Text            =   "32100"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtIP 
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Top             =   2760
         Width           =   1410
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   4320
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Local IP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Local Port:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remote Port:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label lblRemoteIP 
         Caption         =   "Remote IP :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSWinsockLib.Winsock wskChat 
      Index           =   0
      Left            =   4080
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   32100
   End
   Begin VB.TextBox txtReco 
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   5415
   End
   Begin VB.Frame Frame4 
      Caption         =   "Engine Creation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   3135
      Begin VB.OptionButton SharedRC 
         Caption         =   "Shared"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Inproc 
         Caption         =   "Inproc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CheckBox ActivateMic 
      Caption         =   "Activate Mic"
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   720
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.ComboBox SREngines 
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "SREngines"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton ExitBtn 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   3480
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Recognition 
      Caption         =   "Start Recognition"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label CCLabel 
      AutoSize        =   -1  'True
      Caption         =   "Current Grammar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MenuRecognition 
      Caption         =   "&Recognition"
      Begin VB.Menu LoadGrammar 
         Caption         =   "&Load Grammar..."
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
'Author : Troy Drysdale   * Some code taken from Speech API tutorials c. Microsoft
'Date : May 13/02
'Purpose : Demonstrate some cool Speech recognition and use it to create a talking chat
'          Will eventually be made into a C++ program to be used in multiplayer games,
'          as a low bandwidth solution to other communication software.
'          Program may need Microsoft SAPI 5.1 to work.  Download from Microsoft site.
'Contact : td2@post.queensu.ca if you want to contact me with questions or flames.
'=============================================================================

Option Explicit

Public WithEvents RC As SpSharedRecoContext 'The main shared Recognizer Context
Attribute RC.VB_VarHelpID = -1
Public Grammar As ISpeechRecoGrammar        'Command and Control interface
Dim indent As Integer                       'Sets indent level for small output window
Dim fRecoEnabled As Boolean                 'Is recognition enabled
Dim fGrammarLoaded As Boolean               'Is a grammar loaded
Dim RecoResult As ISpeechRecoResult         'Recognition result interface
Dim WithEvents Voice As SpVoice             'The main voice object for talking
Attribute Voice.VB_VarHelpID = -1
Dim m_speakFlags As SpeechVoiceSpeakFlags   'Flags for the voice object
Dim num_Connects As Integer
Const MAX_CONNECTS = 8
Dim bHost As Boolean
Dim bPlayer As Boolean

'Subroutine closes the active connection and resets some flags.
Private Sub cmdDisconnect_Click()
'Clear booleans
    bHost = False
    bPlayer = False
    wskChat(0).Close
'Add a network send to let the host or players know the user is gone hear


End Sub

Private Sub cmdHost_Click()
    
'Check for being a host or player
    If bHost = True Then
        MsgBox "You are already a host!", vbExclamation, "Host Error"
    ElseIf bPlayer = False Then
        wskChat(0).Listen
        bHost = True
    Else
        MsgBox "You must disconnect to do this!", vbExclamation, "Host Error"
    End If
End Sub

Private Sub cmdJoin_Click()
'Check for being a host or player so program doesn't crash
    If bPlayer = True Then
        MsgBox "You are already a player!", vbExclamation, "Join Error"
    ElseIf bHost = False Then
        If txtIP <> "" And txtRemotePort <> "" Then
            wskChat(0).Connect txtIP, txtRemotePort
        End If
    Else
        MsgBox "You are already hosting!", vbExclamation, "Join Error"
    End If
End Sub

Private Sub cmdRepeat_Click()
'Repeats the last captured dictation
    Voice.Speak txtReco.Text
End Sub

Private Sub Form_Load()
'   Set up error handler
    On Error GoTo Err_SAPILoad
'   Initialize globals
    indent = 0
    fRecoEnabled = False
    fGrammarLoaded = False
'   Create the Shared Reco Context by default
    Set RC = New SpSharedRecoContext

'   Load the SR Engines combo box
    Dim Token As ISpeechObjectToken
    For Each Token In RC.Recognizer.GetRecognizers
        SREngines.AddItem Token.GetDescription()
    Next
    SREngines.ListIndex = 0
    
'   Disable combo box for Shared Engine. Also disable other UI that's not initially needed.
    SREngines.Enabled = False
    ActivateMic.Enabled = False
        

'   Create grammar objects
    LoadGrammarObj
   
'   Attempt to load the default .xml file and set the RuleId State to Inactive until
'   the user starts recognition.
    LoadDefaultCnCGrammar
    
'   Voice speaking, creates new objects and sets the flags
    Set Voice = New SpVoice
    m_speakFlags = SVSFlagsAsync Or SVSFPurgeBeforeSpeak Or SVSFIsXML
    
'   Winsock server control set up.
    txtLocalIP = wskChat(0).LocalIP
    txtLocalPort = wskChat(0).LocalPort
    num_Connects = 1
    bHost = False
    bPlayer = False
    Exit Sub
    
Err_SAPILoad:
    MsgBox "Error loading SAPI objects! Please make sure SAPI5.1 is correctly installed.", vbCritical
    Exit_Click
    Exit Sub
End Sub

'   This subroutine creates the Grammar object and sets the states to inactive
'   until the user is ready to begin recognition.
Private Sub LoadGrammarObj()
    Set Grammar = RC.CreateGrammar(1)
    
'   Load Dictation but set it to Inactive until user starts to dictate
    Grammar.DictationLoad "", SLOStatic
    Grammar.DictationSetState SGDSInactive
End Sub

'   This subroutine attempts to load the default English .xml file. It will prompt the
'   user to load a valid .xml file if it cannot find sol.xml in either of the 2
'   specified paths.
Private Sub LoadDefaultCnCGrammar()
'   First load attempt
    On Error Resume Next
    Grammar.CmdLoadFromFile "sol.xml", SLODynamic
    
'   Second load attempt
    If Err Then
    On Error GoTo Err_CFGLoad
        Grammar.CmdLoadFromFile App.Path & "sol.xml", SLODynamic
    End If
    
'   Set rule state to inactive until user clicks Recognition button
    Grammar.CmdSetRuleIdState 0, SGDSInactive
    
'   Set the Label to indictate which .xml file is loaded.
    CCLabel.Caption = "Current C+C Grammar: sol.xml"
    
    fGrammarLoaded = True
    
    Exit Sub
    
Err_CFGLoad:
    fGrammarLoaded = False
    CCLabel.Caption = "Current C+C Grammar: NULL"
    Exit Sub
End Sub

'   This subroutine calls the Common File Dialog control which is inserted into
'   the form to select a .xml grammar file.
Private Sub LoadGrammar_Click()
    ComDlg.CancelError = True
    On Error GoTo Cancel
    ComDlg.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn
    ComDlg.DialogTitle = "Open XML File"
    ComDlg.Filter = "All Files (*.*)|*.*|XML Files " & "(*.xml)|*.xml"
    ComDlg.FilterIndex = 2
    ComDlg.ShowOpen
        
'   Inactivate the grammar and associate a new .xml file with the grammar.
    On Error GoTo Err_XMLLoad
    Grammar.CmdLoadFromFile ComDlg.FileName, SLODynamic
    Grammar.CmdSetRuleIdState 0, SGDSInactive
        
'   Then reactivate the grammar if it was currently enabled
    If fRecoEnabled Then
        Grammar.CmdSetRuleIdState 0, SGDSActive
    End If
    
'   Set the Label to indictate which .xml file is loaded
    CCLabel.Caption = "Current C+C Grammar: " + ComDlg.FileTitle
    
    fGrammarLoaded = True
    Exit Sub
    
Err_XMLLoad:
    fGrammarLoaded = False
    MsgBox "Invalid .xml file. Please load a valid .xml grammar file.", vbOKOnly
    Exit Sub
    
Cancel:
    Exit Sub
End Sub

'   Activates or Deactivates either Command and Control or Dictation based on the
'   current state of the Recognition button.
Private Sub Recognition_Click()
    ActivateMic.Value = Checked
    
    On Error GoTo ErrorHandle

'   First make sure a valid .xml file is loaded if the user is selecting C&C
    If fGrammarLoaded Then
    
    '   If recognition is currently not enabled, enable it.
        If Not fRecoEnabled Then
            Grammar.DictationSetState SGDSActive
            fRecoEnabled = True
    '       Update caption on button.
            Recognition.Caption = "Stop Recognition"
    '       Allow user to activate/deactivate mute
            ActivateMic.Enabled = True
    '       Disable radio buttons and engines combo while recognizing so user doesn't
    '       switch modes during recognition.
            SREngines.Enabled = False
            SharedRC.Enabled = False
            Inproc.Enabled = False
    '   If recognition is currently enabled, disable it.
        Else
            Grammar.DictationSetState SGDSInactive
            fRecoEnabled = False
    '       Update caption on button.
            Recognition.Caption = "Start Recognition"
    '       Disallow user to activate/deactivate mute
            ActivateMic.Enabled = False
    '       Reenable radio buttons while not recognizing so user can switch modes
            
    '       Allow engine selection if the inproc button is selected.
            If Inproc.Value Then
                SREngines.Enabled = True
            End If
            
            SharedRC.Enabled = True
            Inproc.Enabled = True
        End If
    End If
    Exit Sub
    
ErrorHandle:
    MsgBox "Failed to activate the grammar. It is possible that your audio device is used by other application.", vbOKOnly
End Sub

'   False Recognition event handler
Private Sub RC_FalseRecognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal Result As SpeechLib.ISpeechRecoResult)
    Set RecoResult = Result
    Reco_Text
End Sub
'   Hypothesis event handler
Private Sub RC_Hypothesis(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal Result As SpeechLib.ISpeechRecoResult)
    Set RecoResult = Result
    Reco_Text
End Sub
'   Recognition event handler
Private Sub RC_Recognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal RecognitionType As SpeechLib.SpeechRecognitionType, ByVal Result As SpeechLib.ISpeechRecoResult)

    Static i As Integer
    Set RecoResult = Result
    Reco_Text
    Recognition_Click
    'dssSay.Speak txtReco.Text
    Voice.Speak txtReco.Text
    Send_Text
End Sub
Private Sub Reco_Text()
    txtReco.Text = RecoResult.PhraseInfo.GetText()
End Sub

'   The following 2 subroutines destroy/create Inproc and Shared RecoContext's depending
'   on what the user has checked.

'   This subroutine destroys the Inproc RecoContext and creates and Shared RecoContext
Private Sub SharedRC_Click()
'   Destroy the Inproc RecoContext
    Set RC = Nothing
    
'   Create the Shared RecoContext
    Set RC = New SpSharedRecoContext
    
'   Call the InitEventInterestCheckBoxes subroutine which uses the SR engine
'   default event interests to initialize the event interest checkboxes.

'   Create grammar objects
    LoadGrammarObj
   
'   Attempt to load the default .xml file and set the RuleId State to Inactive until
'   the user starts recognition.
    LoadDefaultCnCGrammar
    
'   Disable the engine selection drop down box and reset to the default shared engine.
    SREngines.ListIndex = 0
    SREngines.Enabled = False
End Sub
'   This subroutine destroys the Shared RecoContext and creates and Inproc RecoContext
Private Sub Inproc_Click()
    Dim Recognizer As ISpeechRecognizer
    
'   Destroy Shared RecoContext
    Set RC = Nothing
    
'   Create Inproc Recognizer which we will use to create the Inproc RecoContext.
    Set Recognizer = New SpInprocRecognizer

'   To create an Inproc RecoContext we must set an Audio Input. To do this we create
'   an SpObjectTokenCategory object with the category of AudioIn. This object enumerates
'   the registry to see what types of audio input devices are available.
    Dim ObjectTokenCat As ISpeechObjectTokenCategory
    Set ObjectTokenCat = New SpObjectTokenCategory
    ObjectTokenCat.SetId SpeechCategoryAudioIn

'   Set the default AudioInput device which is typically the first item and is usually
'   the microphone.
    Set Recognizer.AudioInput = ObjectTokenCat.EnumerateTokens.Item(0)
    
'   Set the Recognizer to the one selected in the drop down box.
    Set Recognizer.Recognizer = Recognizer.GetRecognizers().Item(SREngines.ListIndex)

'   Now go ahead and actually create the Inproc RecoContext.
'   Note - in VB even though the global "RC" object is declaired as a
'   SpSharedRecoContext, it is still possible to set it to a SpInprocRecoContext.
    Set RC = Recognizer.CreateRecoContext
    
'   Create grammar objects
    LoadGrammarObj
   
'   Attempt to load the default .xml file and set the RuleId State to Inactive until
'   the user starts recognition.
    LoadDefaultCnCGrammar
    
'   Enable the engine selection drop down box
    SREngines.Enabled = True
End Sub

'   The remaining subroutines handle simple UI and exiting.

'   This subroutine activates/deactivates the microphone.
Private Sub ActivateMic_Click()
    If ActivateMic.Value = Checked Then
            Grammar.DictationSetState SGDSActive
    Else
            Grammar.DictationSetState SGDSInactive
    End If
End Sub
'   This subroutine changes the SR Engine to the selected one
Private Sub SREngines_Click()
'   This subroutine can be called when you update the listindex of SREngines in the form load subroutine
    If Inproc.Value Then
        Inproc_Click
    End If
End Sub
'   About box
Private Sub About_Click()
    MsgBox "Modified By Troy Drysdale" & Chr(13) & "Some code taken from Microsoft. All rights reserved to them." & Chr(13) & "Later versions will be totally rewritten", vbInformation, "About RecoVB"
End Sub
'Unload the open connections...
Private Sub ExitBtn_Click()
    Dim i As Integer

    Set Voice = Nothing
    For i = 0 To wskChat.Count - 1
        wskChat(i).Close
    Next i
    Unload Form1
End Sub
Private Sub Exit_Click()
    Dim i As Integer

    Set Voice = Nothing
    For i = 0 To wskChat.Count - 1
        wskChat(1).Close
    Next i
    Unload Form1
End Sub

Private Sub wskChat_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    
    If num_Connects + 1 < MAX_CONNECTS Then
    'Accept the request, load another socket for the connection
      Load wskChat(num_Connects)
      wskChat(num_Connects).LocalPort = wskChat(num_Connects - 1).LocalPort + 1
      wskChat(num_Connects).Accept requestID
      wskChat(num_Connects).SendData "1:OK"
      num_Connects = num_Connects + 1
    End If
        
End Sub

Private Sub wskChat_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim in_Data As String
    Dim i As Integer
    Dim cmdList As Variant
    Dim out_Data As String
    wskChat(Index).GetData in_Data
    
    cmdList = Split(in_Data, ":")
'Element 0 holds the specified command to be done.
    Select Case cmdList(0)
    Case "1":
            'Connection acceptance
            wskChat(0).SendData "2:" & txtHandle
            bPlayer = True
    Case "2":
            'recieved the user name... sending back list
            lstConnected.AddItem cmdList(1), Index - 1
            out_Data = ""
            out_Data = "3:" & txtHandle.Text
            For i = 0 To lstConnected.ListCount
                If i <> Index - 1 Then
                    out_Data = out_Data & lstConnected.List(i) & ":"
                End If
            Next i
            wskChat(Index).SendData out_Data
    Case "3":
            'recieved list from server.  add to table
            i = 1
            On Error GoTo CASE3_BREAK
            Do While cmdList(i) <> Null
                lstConnected.AddItem cmdList(i), i - 1
            Loop
    Case "4":
            'Received text from server
            txtReco.Text = cmdList(1)
            Voice.Speak cmdList(1)
    Case "5":
            'Received text from player... re-route to all.
            txtReco.Text = cmdList(1)
            Voice.Speak cmdList(1)
            Send_Text
    End Select
CASE3_BREAK:
    
End Sub

Private Sub wskChat_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Error : " & Description
End Sub
Private Sub Send_Text()
Dim i As Integer
If bHost = True Then
    For i = 1 To wskChat.Count - 1
        wskChat(i).SendData "4:" & txtReco.Text
    Next i
ElseIf bPlayer = True Then
    wskChat(0).SendData "5:" & txtReco.Text
End If
End Sub
