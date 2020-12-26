VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmWords 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Audivididici"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmWords.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   11790
   StartUpPosition =   1  'CenterOwner
   Begin Audivididici.RTFUniversalUnicode RTFUser_Answer 
      Height          =   2055
      Left            =   8160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3625
   End
   Begin VB.ComboBox cmbUser_Answer 
      Height          =   315
      ItemData        =   "frmWords.frx":08CA
      Left            =   7680
      List            =   "frmWords.frx":08D4
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Timer tmrSoundTracker 
      Interval        =   10
      Left            =   480
      Top             =   1560
   End
   Begin VB.Timer tmrPressSpace 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1200
      Top             =   600
   End
   Begin VB.Timer tmrAutoClick 
      Enabled         =   0   'False
      Left            =   720
      Top             =   480
   End
   Begin MCI.MMControl MMCUitspraak 
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdUser_Wrong 
      Caption         =   "&Fout"
      Height          =   615
      Left            =   5280
      Picture         =   "frmWords.frx":08F2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdUser_Good 
      Caption         =   "&Goed"
      Height          =   615
      Left            =   3720
      Picture         =   "frmWords.frx":0C06
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdUser_ShowAnswer 
      Caption         =   "Toon &Antwoord"
      Height          =   615
      Left            =   4320
      Picture         =   "frmWords.frx":0D4E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox txtWord_Question 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"frmWords.frx":0DD2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Gentium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtWord_Answer 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4680
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"frmWords.frx":0E50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6690
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8379
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2461
            MinWidth        =   2469
            Text            =   "Resterend: 0"
            TextSave        =   "Resterend: 0"
            Object.ToolTipText     =   "Resterend: 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Picture         =   "frmWords.frx":0EC9
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Goed: 0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Picture         =   "frmWords.frx":1021
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Fout: 0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "|||||||||||||||||||||||||||||||||||"
            TextSave        =   "|||||||||||||||||||||||||||||||||||"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "18:36"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtUser_Answer 
      Height          =   615
      Left            =   1680
      TabIndex        =   10
      Top             =   5760
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmWords.frx":1345
   End
   Begin VB.Label lblPressSpace 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Druk op de spatiebalk om door te gaan!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Shape shpWord 
      BorderColor     =   &H000080FF&
      Height          =   2055
      Left            =   0
      Top             =   3600
      Width           =   11415
   End
   Begin VB.Image imgWord_Picture 
      Height          =   3495
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type LearnType
    Question As String
    Answer As String
    Picture As StdPicture
    Pronouncement As String
    Good As Integer
    Wrong As Integer
    GoodTotal As Integer
    WrongTotal As Integer
    Done As Boolean
    Language As String
End Type

Private Enum LearnMethod
    LearnMethod_Slideshow = 0
    LearnMethod_Practise = 1
    LearnMethod_Copy = 2
    LearnMethod_Test = 3
    LearnMethod_Pronouncement = 4
End Enum


Private LearnDatabase(999) As LearnType
Private LearnDatabaseCount As Integer
Private CurrentWord As Integer
Private CurrentLearnMethod As LearnMethod
Private GlobalFilePath As String
Private GlobalFileName As String
Private GlobalFileDescription As String
Private GlobalImages As Boolean
Private GlobalSound As Boolean
Private GlobalRepeat As Boolean
Private GlobalGoOn As Boolean
Private GlobalSwitch As Boolean
Private GlobalRandom As Boolean

Sub ShowCurrentWord()
    Dim Goed_Totaal As Integer
    Dim Fout_Totaal As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Top3_Hardest(2) As String
    Dim Top3_Hardest_Total(2) As String
    Dim Top3_Best_Progress(2) As String
    Dim Hardest_Wrong(2) As Integer
    Dim Hardest_Total_Wrong(2) As Integer
    Dim Best_Progress_Ratio(2) As Double
    Dim Language As String
    
    If CurrentWord < 0 Then
        If CurrentLearnMethod = LearnMethod_Slideshow Then
            Unload Me
            Exit Sub
        End If
        Goed_Totaal = 0
        Fout_Totaal = 0
        
        For i = 0 To 2
            Hardest_Wrong(i) = 0
            Hardest_Total_Wrong(i) = 0
            Best_Progress_Ratio(i) = -1
            Top3_Hardest(i) = ""
            Top3_Hardest_Total(i) = ""
            Top3_Best_Progress(i) = ""
        Next i
        
        If LearnDatabaseCount >= 10 Then
            For i = 0 To LearnDatabaseCount - 1
                If LearnDatabase(i).Question <> "" Then
                    Goed_Totaal = Goed_Totaal + IIf(LearnDatabase(i).Wrong <= 0, 1, 0)
                    Fout_Totaal = Fout_Totaal + IIf(LearnDatabase(i).Wrong > 0, 1, 0)
                    
                    If LearnDatabase(i).Wrong > Hardest_Wrong(0) Then
                        Hardest_Wrong(2) = Hardest_Wrong(1)
                        Hardest_Wrong(1) = Hardest_Wrong(0)
                        Hardest_Wrong(0) = LearnDatabase(i).Wrong
                        Top3_Hardest(2) = Top3_Hardest(1)
                        Top3_Hardest(1) = Top3_Hardest(0)
                        Top3_Hardest(0) = LearnDatabase(i).Question
                    ElseIf LearnDatabase(i).Wrong > Hardest_Wrong(1) Then
                        Hardest_Wrong(2) = Hardest_Wrong(1)
                        Hardest_Wrong(1) = LearnDatabase(i).Wrong
                        Top3_Hardest(2) = Top3_Hardest(1)
                        Top3_Hardest(1) = LearnDatabase(i).Question
                    ElseIf LearnDatabase(i).Wrong > Hardest_Wrong(2) Then
                        Hardest_Wrong(2) = LearnDatabase(i).Wrong
                        Top3_Hardest(2) = LearnDatabase(i).Question
                    End If
                    
                    If LearnDatabase(i).WrongTotal > Hardest_Total_Wrong(0) Then
                        Hardest_Total_Wrong(2) = Hardest_Total_Wrong(1)
                        Hardest_Total_Wrong(1) = Hardest_Total_Wrong(0)
                        Hardest_Total_Wrong(0) = LearnDatabase(i).WrongTotal
                        Top3_Hardest_Total(2) = Top3_Hardest_Total(1)
                        Top3_Hardest_Total(1) = Top3_Hardest_Total(0)
                        Top3_Hardest_Total(0) = LearnDatabase(i).Question
                    ElseIf LearnDatabase(i).WrongTotal > Hardest_Total_Wrong(1) Then
                        Hardest_Total_Wrong(2) = Hardest_Total_Wrong(1)
                        Hardest_Total_Wrong(1) = LearnDatabase(i).WrongTotal
                        Top3_Hardest_Total(2) = Top3_Hardest_Total(1)
                        Top3_Hardest_Total(1) = LearnDatabase(i).Question
                    ElseIf LearnDatabase(i).WrongTotal > Hardest_Total_Wrong(2) Then
                        Hardest_Total_Wrong(2) = LearnDatabase(i).WrongTotal
                        Top3_Hardest_Total(2) = LearnDatabase(i).Question
                    End If
                    
                    If LearnDatabase(i).WrongTotal > 0 Then
                        If LearnDatabase(i).Wrong > 0 Then
                            If CDbl(LearnDatabase(i).Wrong) / CDbl(LearnDatabase(i).WrongTotal) < Best_Progress_Ratio(0) Or Best_Progress_Ratio(0) = -1 Then
                                Best_Progress_Ratio(2) = Best_Progress_Ratio(1)
                                Best_Progress_Ratio(1) = Best_Progress_Ratio(0)
                                Best_Progress_Ratio(0) = CDbl(LearnDatabase(i).Wrong) / CDbl(LearnDatabase(i).WrongTotal)
                                Top3_Best_Progress(2) = Top3_Best_Progress(1)
                                Top3_Best_Progress(1) = Top3_Best_Progress(0)
                                Top3_Best_Progress(0) = LearnDatabase(i).Question
                            ElseIf CDbl(LearnDatabase(i).Wrong) / CDbl(LearnDatabase(i).WrongTotal) < Best_Progress_Ratio(1) Or Best_Progress_Ratio(1) = -1 Then
                                Best_Progress_Ratio(2) = Best_Progress_Ratio(1)
                                Best_Progress_Ratio(1) = CDbl(LearnDatabase(i).Wrong) / CDbl(LearnDatabase(i).WrongTotal)
                                Top3_Best_Progress(2) = Top3_Best_Progress(1)
                                Top3_Best_Progress(1) = LearnDatabase(i).Question
                            ElseIf CDbl(LearnDatabase(i).Wrong) / CDbl(LearnDatabase(i).WrongTotal) < Best_Progress_Ratio(2) Or Best_Progress_Ratio(2) = -1 Then
                                Best_Progress_Ratio(2) = CDbl(LearnDatabase(i).Wrong) / CDbl(LearnDatabase(i).WrongTotal)
                                Top3_Best_Progress(2) = LearnDatabase(i).Question
                            End If
                        Else
                            If -CDbl(LearnDatabase(i).WrongTotal) < Best_Progress_Ratio(0) Or Best_Progress_Ratio(0) = -1 Then
                                Best_Progress_Ratio(2) = Best_Progress_Ratio(1)
                                Best_Progress_Ratio(1) = Best_Progress_Ratio(0)
                                Best_Progress_Ratio(0) = -CDbl(LearnDatabase(i).WrongTotal)
                                Top3_Best_Progress(2) = Top3_Best_Progress(1)
                                Top3_Best_Progress(1) = Top3_Best_Progress(0)
                                Top3_Best_Progress(0) = LearnDatabase(i).Question
                            ElseIf -CDbl(LearnDatabase(i).WrongTotal) < Best_Progress_Ratio(1) Or Best_Progress_Ratio(1) = -1 Then
                                Best_Progress_Ratio(2) = Best_Progress_Ratio(1)
                                Best_Progress_Ratio(1) = -CDbl(LearnDatabase(i).WrongTotal)
                                Top3_Best_Progress(2) = Top3_Best_Progress(1)
                                Top3_Best_Progress(1) = LearnDatabase(i).Question
                            ElseIf -CDbl(LearnDatabase(i).WrongTotal) < Best_Progress_Ratio(2) Or Best_Progress_Ratio(2) = -1 Then
                                Best_Progress_Ratio(2) = -CDbl(LearnDatabase(i).WrongTotal)
                                Top3_Best_Progress(2) = LearnDatabase(i).Question
                            End If
                        End If
                    End If
                End If
            Next i
        
            frmResults.SetTop3 Top3_Hardest(0), Top3_Hardest(1), Top3_Hardest(2), Top3_Hardest_Total(0), Top3_Hardest_Total(1), Top3_Hardest_Total(2), Top3_Best_Progress(0), Top3_Best_Progress(1), Top3_Best_Progress(2), True
        Else
            For i = 0 To LearnDatabaseCount - 1
                If LearnDatabase(i).Question <> "" Then
                    Goed_Totaal = Goed_Totaal + IIf(LearnDatabase(i).Wrong <= 0, 1, 0)
                    Fout_Totaal = Fout_Totaal + IIf(LearnDatabase(i).Wrong > 0, 1, 0)
                End If
            Next i
            
            frmResults.SetTop3 "", "", "", "", "", "", "", "", "", False
        End If
        
        If (Goed_Totaal + Fout_Totaal > 0) Then
            frmResults.NewScore Val((Goed_Totaal / (Goed_Totaal + Fout_Totaal)) * 100)
            frmResults.Show vbModal
        End If
    Else
        txtWord_Question.TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}}\viewkind4\uc1\pard\f0\fs38\qc " & LearnDatabase(CurrentWord).Question & "\par}"
        txtWord_Answer.TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}}\viewkind4\uc1\pard\f0\fs32\qc " & LearnDatabase(CurrentWord).Answer & "\par}"
        
        If IsObject(LearnDatabase(CurrentWord).Picture) Then
            Set imgWord_Picture.Picture = LearnDatabase(CurrentWord).Picture
        Else
            Set imgWord_Picture.Picture = Nothing
        End If
        
        If CurrentLearnMethod <> LearnMethod_Pronouncement And LearnDatabase(CurrentWord).Pronouncement <> "" Then
            PlayFile LearnDatabase(CurrentWord).Pronouncement
        End If
        
        Form_Resize
        
        Select Case CurrentLearnMethod
            Case LearnMethod_Slideshow
                txtWord_Answer.Visible = True
                cmdUser_ShowAnswer.Caption = "&Volgende"
                cmdUser_ShowAnswer.Visible = True
                cmdUser_ShowAnswer.Default = True
                cmdUser_Good.Visible = False
                cmdUser_Wrong.Visible = False
                txtUser_Answer.Visible = False
                If tmrAutoClick.Interval <> 0 Then
                    tmrAutoClick.Enabled = True
                Else
                    tmrAutoClick.Enabled = False
                End If
                lblPressSpace.Visible = False
                tmrPressSpace.Enabled = False
                cmbUser_Answer.Visible = False
                RTFUser_Answer.Visible = False
            Case LearnMethod_Practise
                txtWord_Answer.Visible = False
                cmdUser_ShowAnswer.Caption = "Toon &antwoord"
                cmdUser_ShowAnswer.Visible = True
                cmdUser_ShowAnswer.Default = True
                cmdUser_Good.Caption = "&Goed"
                cmdUser_Good.Visible = False
                cmdUser_Wrong.Caption = "&Fout"
                cmdUser_Wrong.Visible = False
                txtUser_Answer.Visible = False
                If tmrAutoClick.Interval <> 0 Then
                    tmrAutoClick.Enabled = True
                Else
                    tmrAutoClick.Enabled = False
                End If
                lblPressSpace.Visible = False
                tmrPressSpace.Enabled = False
                cmbUser_Answer.Visible = False
                RTFUser_Answer.Visible = False
            Case LearnMethod_Copy
                txtWord_Answer.Visible = True
                cmdUser_ShowAnswer.Visible = False
                cmdUser_Good.Visible = False
                cmdUser_Wrong.Visible = False
                txtUser_Answer.Visible = True
                txtUser_Answer.Locked = False
                tmrAutoClick.Enabled = False
                txtUser_Answer.TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}}\viewkind4\uc1\pard\f0\fs32\qc  \par}"
                txtUser_Answer.BackColor = &HFFFFFF
                lblPressSpace.Visible = False
                tmrPressSpace.Enabled = False
                cmbUser_Answer.Visible = True
                RTFUser_Answer.Visible = True
                Language = GetLanguage(CurrentWord, False)
                If Language = "L" Then
                    cmbUser_Answer.ListIndex = 0
                ElseIf Language = "G" Then
                    cmbUser_Answer.ListIndex = 1
                End If
                If Me.Visible Then txtUser_Answer.SetFocus
            Case LearnMethod_Test
                txtWord_Answer.Visible = False
                cmdUser_ShowAnswer.Visible = False
                cmdUser_Good.Visible = False
                cmdUser_Wrong.Visible = False
                txtUser_Answer.Visible = True
                txtUser_Answer.Locked = False
                tmrAutoClick.Enabled = False
                txtUser_Answer.TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}}\viewkind4\uc1\pard\f0\fs32\qc  \par}"
                txtUser_Answer.BackColor = &HFFFFFF
                lblPressSpace.Visible = False
                tmrPressSpace.Enabled = False
                cmbUser_Answer.Visible = True
                RTFUser_Answer.Visible = True
                Language = GetLanguage(CurrentWord, False)
                If Language = "L" Then
                    cmbUser_Answer.ListIndex = 0
                ElseIf Language = "G" Then
                    cmbUser_Answer.ListIndex = 1
                End If
                If Me.Visible Then txtUser_Answer.SetFocus
            Case LearnMethod_Pronouncement
                txtWord_Answer.Visible = True
                cmdUser_ShowAnswer.Caption = "Beluister &antwoord"
                cmdUser_ShowAnswer.Visible = True
                cmdUser_ShowAnswer.Default = True
                cmdUser_Good.Caption = "&Goed"
                cmdUser_Good.Visible = False
                cmdUser_Wrong.Caption = "&Fout"
                cmdUser_Wrong.Visible = False
                txtUser_Answer.Visible = False
                If tmrAutoClick.Interval <> 0 Then
                    tmrAutoClick.Enabled = True
                Else
                    tmrAutoClick.Enabled = False
                End If
                lblPressSpace.Visible = False
                tmrPressSpace.Enabled = False
                cmbUser_Answer.Visible = False
                RTFUser_Answer.Visible = False
        End Select
        
        Refresh_StatusBar
    End If
End Sub

Public Function GetRandomWord()
    Dim i As Integer
    Dim WordLeft As Boolean
    Randomize
    
    i = 0
    WordLeft = False
    Do While LearnDatabase(i).Question <> ""
            If LearnDatabase(i).Done = False Then WordLeft = True
            i = i + 1
    Loop
    
    If Not WordLeft Then
        GetRandomWord = -1
        Exit Function
    End If
    
    If GlobalRandom Then
        Do
            i = CInt(Rnd() * (LearnDatabaseCount - 1))
        Loop While LearnDatabase(i).Done = True Or LearnDatabase(i).Question = ""
    Else
        i = 0
        Do While LearnDatabase(i).Question <> "" And LearnDatabase(i).Done = True
                i = i + 1
        Loop
    End If
    
    GetRandomWord = i
End Function

Sub PlayFile(FileName As String)
    If FExists(AppPath & "AVDTemp\" & FileName) And GlobalSound Then
        MMCUitspraak.Command = "stop"
        MMCUitspraak.Command = "close"
        MMCUitspraak.FileName = AppPath & "AVDTemp\" & FileName
        MMCUitspraak.Command = "open"
        MMCUitspraak.Command = "play"
    End If
End Sub

Sub ClearDatabase()
    Dim i As Integer
    LearnDatabaseCount = 0
    For i = 0 To 999
        LearnDatabase(i).Question = ""
        LearnDatabase(i).Answer = ""
        'LearnDatabase(i).Picture = Empty
        LearnDatabase(i).Pronouncement = ""
        LearnDatabase(i).Good = 0
        LearnDatabase(i).Wrong = 0
        LearnDatabase(i).GoodTotal = 0
        LearnDatabase(i).WrongTotal = 0
        LearnDatabase(i).Done = False
    Next i
End Sub

Function LoadFile(FilePath As String, FileName As String, FileDescription As String, ClickSpeed As Integer, UseImages As Boolean, UseSound As Boolean, Repeat As Boolean, GoOn As Boolean, Switch As Boolean, Random As Boolean) As Boolean
    Dim MyFreeFile As Integer
    Dim MyLine As String
    Dim MyWordNumber As Integer
    Dim TempMax As Integer
    Dim RetCode As Long
    
    MyFreeFile = FreeFile()
    MyWordNumber = 0
    
    If Right$(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    If Not FExists(FilePath & FileName) Then
        MsgBox "Bestand niet gevonden!", vbCritical, "Audivididici"
        LoadFile = False
        Exit Function
    End If
    
    Load frmSplash
    frmSplash.SetMaximumAndValue 1, 0
    frmSplash.Show
    
    tmrAutoClick.Enabled = False
    tmrPressSpace.Enabled = False
    tmrSoundTracker.Enabled = False
    
    GlobalFilePath = FilePath
    GlobalFileName = FileName
    GlobalFileDescription = FileDescription
    GlobalImages = UseImages
    GlobalSound = UseSound
    GlobalRepeat = Repeat
    GlobalGoOn = GoOn
    GlobalSwitch = Switch
    GlobalRandom = Random
    
    Me.Caption = GlobalFileDescription & " - Audivididici"
    
    ClearDatabase
    
    Load frmResults
    frmResults.ClearScores
    
    MMCUitspraak.Command = "close"
    DelDir AppPath & "AVDTemp\"
    Dir AppPath
    
    '-- Init Global Message Variables
    uZipInfo = ""
    uZipNumber = 0   ' Holds The Number Of Zip Files
    
    '-- Select UNZIP32.DLL Options - Change As Required!
    uPromptOverWrite = 0  ' 1 = Prompt To Overwrite
    uOverWriteFiles = 1   ' 1 = Always Overwrite Files
    uDisplayComment = 0   ' 1 = Display comment ONLY!!!
    
    '-- Change The Next Line To Do The Actual Unzip!
    uExtractList = 0       ' 1 = List Contents Of Zip 0 = Extract
    uHonorDirectories = 0  ' 1 = Honour Zip Directories
    
    '-- Select Filenames If Required
    '-- Or Just Select All Files
    uZipNames.uzFiles(0) = vbNullString
    uNumberFiles = 0
    
    '-- Select Filenames To Exclude From Processing
    ' Note UNIX convention!
    '   vbxnames.s(0) = "VBSYX/VBSYX.MID"
    '   vbxnames.s(1) = "VBSYX/VBSYX.SYX"
    '   numx = 2
    
    '-- Or Just Select All Files
    uExcludeNames.uzFiles(0) = vbNullString
    uNumberXFiles = 0
    
    '-- Change The Next 2 Lines As Required!
    '-- These Should Point To Your Directory
    uZipFileName = GlobalFilePath & GlobalFileName
    uExtractDir = AppPath & "AVDTemp\"
    
    '-- Let's Go And Unzip Them!
    RetCode = VBUnZip32
    
    If RetCode <> 0 Then
        MsgBox "Er is een foutmelding opgetreden: UNZIP_" & Str$(RetCode), vbCritical, "Audivididici Creator"
        Unload frmSplash
        LoadFile = False
        Unload Me
        Exit Function
    End If
    
    Open AppPath & "AVDTemp\info.txt" For Input As #MyFreeFile
        Line Input #MyFreeFile, MyLine
        TempMax = Val(MyLine)
        If TempMax > 0 Then frmSplash.SetMaximumAndValue TempMax + 1, 1
        
        Do While Not EOF(MyFreeFile)
          
            Line Input #MyFreeFile, MyLine
            MyLine = Trim$(MyLine)
            If Not GlobalSwitch Then
                LearnDatabase(MyWordNumber).Question = MyLine
            Else
                LearnDatabase(MyWordNumber).Answer = MyLine
            End If
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            
            Line Input #MyFreeFile, MyLine
            If Not GlobalSwitch Then
                LearnDatabase(MyWordNumber).Answer = MyLine
            Else
                LearnDatabase(MyWordNumber).Question = MyLine
            End If
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            
            Line Input #MyFreeFile, MyLine
            MyLine = Trim$(MyLine)
            If FExists(AppPath & "AVDTemp\" & MyLine) Then
                Set LearnDatabase(MyWordNumber).Picture = StdFunctions.LoadPicture(AppPath & "AVDTemp\" & MyLine)
            Else
                Set LearnDatabase(MyWordNumber).Picture = Nothing
            End If
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            
            Line Input #MyFreeFile, MyLine
            MyLine = Trim$(MyLine)
            If FExists(AppPath & "AVDTemp\" & MyLine) Then
                LearnDatabase(MyWordNumber).Pronouncement = MyLine
            End If
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            
            Line Input #MyFreeFile, MyLine
            MyLine = Trim$(MyLine)
            If Not GlobalSwitch Then
                LearnDatabase(MyWordNumber).Language = MyLine
            Else
                LearnDatabase(MyWordNumber).Language = StrReverse(MyLine)
            End If

DoEinde:
            
            MyWordNumber = MyWordNumber + 1
            
            frmSplash.SetValue MyWordNumber + 1
            DoEvents
        Loop
    Close #MyFreeFile
    
    LearnDatabaseCount = MyWordNumber
    
    If ClickSpeed < 16 Then
        tmrAutoClick.Interval = ClickSpeed * 1000
    Else
        tmrAutoClick.Interval = 0
    End If
    
    tmrSoundTracker.Enabled = True
    
    CurrentWord = GetRandomWord
    If CurrentWord < 0 Then
        LoadFile = False
    Else
        ShowCurrentWord
        LoadFile = True
    End If
    
    Unload frmSplash
End Function

Private Sub cmbUser_Answer_Click()
    RTFUser_Answer.SetKeyboard cmbUser_Answer.ListIndex
End Sub

Private Sub cmdUser_Good_Click()
    LearnDatabase(CurrentWord).Done = True
    LearnDatabase(CurrentWord).Good = LearnDatabase(CurrentWord).Good + 1
    LearnDatabase(CurrentWord).GoodTotal = LearnDatabase(CurrentWord).GoodTotal + 1
    CurrentWord = GetRandomWord
    ShowCurrentWord
End Sub

Private Sub cmdUser_Wrong_Click()
    LearnDatabase(CurrentWord).Done = IIf(GlobalRepeat, False, True)
    LearnDatabase(CurrentWord).Wrong = LearnDatabase(CurrentWord).Wrong + 1
    LearnDatabase(CurrentWord).WrongTotal = LearnDatabase(CurrentWord).WrongTotal + 1
    CurrentWord = GetRandomWord
    ShowCurrentWord
End Sub

Private Sub cmdUser_ShowAnswer_Click()
    tmrAutoClick.Enabled = False
    
    If CurrentWord >= 0 And CurrentWord < LearnDatabaseCount Then
        Select Case CurrentLearnMethod
            Case LearnMethod_Slideshow
                LearnDatabase(CurrentWord).Done = True
                CurrentWord = GetRandomWord
                ShowCurrentWord
            Case LearnMethod_Practise
                txtWord_Answer.Visible = True
                cmdUser_ShowAnswer.Caption = "Toon &antwoord"
                cmdUser_ShowAnswer.Visible = False
                cmdUser_ShowAnswer.Default = False
                cmdUser_Good.Caption = "&Goed"
                cmdUser_Good.Visible = True
                cmdUser_Wrong.Caption = "&Fout"
                cmdUser_Wrong.Visible = True
                txtUser_Answer.Visible = False
            Case LearnMethod_Pronouncement
                txtWord_Answer.Visible = True
                cmdUser_ShowAnswer.Caption = "Beluister &antwoord"
                cmdUser_ShowAnswer.Visible = False
                cmdUser_ShowAnswer.Default = False
                cmdUser_Good.Caption = "&Goed"
                cmdUser_Good.Visible = True
                cmdUser_Wrong.Caption = "&Fout"
                cmdUser_Wrong.Visible = True
                txtUser_Answer.Visible = False
                If LearnDatabase(CurrentWord).Pronouncement <> "" Then
                    PlayFile LearnDatabase(CurrentWord).Pronouncement
                End If
        End Select
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyA
            If (CurrentLearnMethod = LearnMethod_Pronouncement Or CurrentLearnMethod = LearnMethod_Practise) And cmdUser_Good.Visible = False And cmdUser_Wrong.Visible = False Then
                cmdUser_ShowAnswer_Click
            End If
        Case vbKeyV
            If CurrentLearnMethod = LearnMethod_Slideshow Then
                cmdUser_ShowAnswer_Click
            End If
        Case vbKeyF
            If CurrentLearnMethod = LearnMethod_Pronouncement Or CurrentLearnMethod = LearnMethod_Practise And cmdUser_ShowAnswer.Visible = False Then
                cmdUser_Wrong_Click
            End If
        Case vbKeyG
            If CurrentLearnMethod = LearnMethod_Pronouncement Or CurrentLearnMethod = LearnMethod_Practise And cmdUser_ShowAnswer.Visible = False Then
                cmdUser_Good_Click
            End If
    End Select
End Sub

Private Sub Form_Load()
    StatusBar.Panels(3).Picture = cmdUser_Good.Picture
    StatusBar.Panels(4).Picture = cmdUser_Wrong.Picture
    RTFUser_Answer.HideMenu
    cmbUser_Answer.ListIndex = 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        If Me.Width < 11910 Then Me.Width = 11910
        If Me.Height < 7455 Then Me.Height = 7455
        
        With imgWord_Picture
            If .Picture.Width > 0 And .Picture.Height > 0 And GlobalImages Then
                .Visible = True
                .Height = Me.Height - 4600
                .Width = (.Picture.Width / .Picture.Height) * .Height
                 
                If .Width > Me.Width - 1000 Then
                    .Width = Me.Width - 1000
                    .Height = (.Picture.Height / .Picture.Width) * .Width
                End If
                
                .Top = (Me.Height - 4300 - 200) / 2 + 100 - .Height / 2
                .Left = Me.Width / 2 - .Width / 2
            End If
            If Not GlobalImages Then
                .Visible = False
            End If
        End With
        
        If GlobalImages Then
            shpWord.Top = Me.Height - 4300
        Else
            shpWord.Top = Me.Height / 2
        End If
        shpWord.Left = Me.Width / 2 - shpWord.Width / 2 - 100
        lblPressSpace.Top = shpWord.Top + shpWord.Height + 10
        lblPressSpace.Left = Me.Width / 2 - lblPressSpace.Width / 2
        txtWord_Question.Left = shpWord.Left + 120
        txtWord_Question.Top = shpWord.Top + 120
        txtWord_Answer.Left = shpWord.Left + 120
        txtWord_Answer.Top = txtWord_Question.Top + txtWord_Question.Height + 120
        txtUser_Answer.Top = shpWord.Top + shpWord.Height + 330
        txtUser_Answer.Left = Me.Width / 2 - txtUser_Answer.Width / 2
        RTFUser_Answer.Top = txtUser_Answer.Top - 650
        RTFUser_Answer.Left = txtUser_Answer.Left + txtUser_Answer.Width - 680
        cmbUser_Answer.Top = txtUser_Answer.Top + txtUser_Answer.Height + 200
        cmbUser_Answer.Left = txtUser_Answer.Left + txtUser_Answer.Width - cmbUser_Answer.Width + 530
        cmdUser_ShowAnswer.Top = txtUser_Answer.Top
        cmdUser_ShowAnswer.Left = Me.Width / 2 - cmdUser_ShowAnswer.Width / 2
        cmdUser_Good.Top = txtUser_Answer.Top
        cmdUser_Good.Left = Me.Width / 2 - cmdUser_Good.Width - 100
        cmdUser_Wrong.Top = txtUser_Answer.Top
        cmdUser_Wrong.Left = Me.Width / 2 + 100
    End If
End Sub

Function Streepjes(Procent As Integer) As String
    Dim AantalStreepjes As Integer
    Dim StrStreepjes As String
    AantalStreepjes = Int((Procent / 100) * 35)
    StrStreepjes = ""
    Dim i As Integer
    For i = 1 To AantalStreepjes
        StrStreepjes = StrStreepjes + "|"
    Next i
    Streepjes = StrStreepjes
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload frmResults
    MMCUitspraak.Command = "close"
    DelDir AppPath & "AVDTemp\"
    Dir AppPath
    frmMain.cmdEdit.Enabled = True
    frmMain.cmdStart.Enabled = True
    If Not Demo Then
        frmMain.cmdAdvanced.Enabled = True
    End If
End Sub

Private Sub RTFUser_Answer_ChangeRTF(SelRTF As String)
    DoEvents
    txtUser_Answer.SelRTF = SelRTF
    DoEvents
End Sub

Private Sub RTFUser_Answer_CheckSigma()
    Dim OldSelStart As Integer
    Dim OldSelLength As Integer
    
    If Len(txtUser_Answer.Text) > 0 And txtUser_Answer.SelStart > 0 Then
        OldSelStart = txtUser_Answer.SelStart
        OldSelLength = txtUser_Answer.SelLength
        txtUser_Answer.SelStart = txtUser_Answer.SelStart - 1
        txtUser_Answer.SelLength = 1
        RTFUser_Answer.RTFValidate txtUser_Answer.SelRTF, False
        txtUser_Answer.SelStart = OldSelStart
        txtUser_Answer.SelLength = OldSelLength
    End If
End Sub

Private Sub RTFUser_Answer_Click()
    RTFUser_Answer.ShowMenu 0, 0
End Sub

Private Sub RTFUser_Answer_RTFSetFocus()
    On Error Resume Next
    txtUser_Answer.SetFocus
End Sub

Private Sub tmrSoundTracker_Timer()
    Dim PositieProcent As Integer
    If MMCUitspraak.Position > 0 And MMCUitspraak.Length > 0 And MMCUitspraak.Length <> MMCUitspraak.Position Then
        PositieProcent = Int(100 / (MMCUitspraak.Length / MMCUitspraak.Position))
        StatusBar.Panels(5).Text = Streepjes(PositieProcent)
    Else
        StatusBar.Panels(5).Text = ""
    End If
End Sub

Sub Refresh_StatusBar()
    Dim iRemaining, iGoed, iFout, i As Integer
    With StatusBar
        .Panels(1).Text = GlobalFileDescription
        .Panels(1).ToolTipText = GlobalFileDescription
        
        iRemaining = 0
        For i = 0 To LearnDatabaseCount - 1
            If LearnDatabase(i).Question <> "" Then
                If LearnDatabase(i).Done = False Then iRemaining = iRemaining + 1
                iGoed = iGoed + LearnDatabase(i).Good
                iFout = iFout + LearnDatabase(i).Wrong
            End If
        Next i
        
        .Panels(2).Text = "Resterend: " & Trim$(Str(iRemaining))
        .Panels(2).ToolTipText = "Resterend: " & Trim$(Str(iRemaining))
        .Panels(3).Text = Trim$(Str(iGoed))
        .Panels(3).ToolTipText = "Goed: " & Trim$(Str(iGoed))
        .Panels(4).Text = Trim$(Str(iFout))
        .Panels(4).ToolTipText = "Fout: " & Trim$(Str(iFout))
    End With
End Sub

Private Sub tmrAutoClick_Timer()
    cmdUser_ShowAnswer_Click
End Sub

Private Sub tmrPressSpace_Timer()
    lblPressSpace.Visible = True
End Sub

Private Sub txtUser_Answer_Change()
    Dim TempSelStart As Integer
    Dim TempSelLength As Integer
    
    TempSelStart = txtUser_Answer.SelStart
    TempSelLength = txtUser_Answer.SelLength
    txtUser_Answer.SelStart = 0
    txtUser_Answer.SelLength = Len(txtUser_Answer.Text)
    txtUser_Answer.SelFontName = "Gentium"
    txtUser_Answer.SelFontSize = 18
    txtUser_Answer.SelAlignment = rtfCenter
    txtUser_Answer.SelStart = TempSelStart
    txtUser_Answer.SelLength = TempSelLength
End Sub

Private Sub txtUser_Answer_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim SpacePosition As Integer
    Dim UPosition As Integer
    Dim UserRTF As String
    Dim WordRTF As String
    Dim AnswerGood As Boolean
    
    If KeyAscii = vbKeyReturn And txtUser_Answer.BackColor = &HFFFFFF Then
        txtWord_Answer.Visible = True
        tmrPressSpace.Enabled = True
        txtUser_Answer.Locked = True
        
        AnswerGood = False
        
        If Len(txtUser_Answer.Text) > 0 And Len(txtWord_Answer.Text) > 0 Then
            RTFUser_Answer_CheckSigma
            
            AnswerGood = True
            
            If Trim(Replace(Replace(Replace(txtUser_Answer.Text, vbCr, ""), vbLf, ""), "?", "")) <> Trim(Replace(Replace(txtWord_Answer.Text, vbCr, ""), vbLf, "")) Then
                i = -1
                j = -1
                
                Do
                    If i < Len(txtUser_Answer.Text) - 1 Then
                        Do
                            i = i + 1
                            txtUser_Answer.SelStart = i
                            txtUser_Answer.SelLength = 1
                        Loop While txtUser_Answer.SelText = " "
                    Else
                        i = i + 1
                        txtUser_Answer.SelLength = 0
                        txtUser_Answer.SelStart = Len(txtUser_Answer.Text)
                    End If
                        
                    If j < Len(txtUser_Answer.Text) - 1 Then
                        Do
                            j = j + 1
                            txtWord_Answer.SelStart = j
                            txtWord_Answer.SelLength = 1
                        Loop While txtWord_Answer.SelText = " "
                    Else
                        j = j + 1
                        txtWord_Answer.SelLength = 0
                        txtWord_Answer.SelStart = Len(txtWord_Answer.Text)
                    End If
                    
                    SpacePosition = InStrRev(txtUser_Answer.SelRTF, " ")
                    UPosition = InStrRev(txtUser_Answer.SelRTF, "\")
                    
                    If UPosition > SpacePosition Then
                        UserRTF = Right(txtUser_Answer.SelRTF, Len(txtUser_Answer.SelRTF) - UPosition + 1)
                    Else
                        UserRTF = Right(txtUser_Answer.SelRTF, Len(txtUser_Answer.SelRTF) - SpacePosition)
                    End If
                    
                    SpacePosition = InStrRev(txtWord_Answer.SelRTF, " ")
                    UPosition = InStrRev(txtWord_Answer.SelRTF, "\")
                    
                    If UPosition > SpacePosition Then
                        WordRTF = Right(txtWord_Answer.SelRTF, Len(txtWord_Answer.SelRTF) - UPosition + 1)
                    Else
                        WordRTF = Right(txtWord_Answer.SelRTF, Len(txtWord_Answer.SelRTF) - SpacePosition)
                    End If
                    
                    If UserRTF <> WordRTF Then
                        AnswerGood = False
                    End If
                Loop While i < Len(txtUser_Answer.Text) - 1 Or j < Len(txtWord_Answer.Text) - 1
            End If
        End If
        
        txtWord_Answer.SelLength = 0
        txtWord_Answer.SelStart = 0
        txtUser_Answer.SelLength = 0
        txtUser_Answer.SelStart = Len(txtUser_Answer.Text)
        
        If AnswerGood Then
            LearnDatabase(CurrentWord).Done = True
            LearnDatabase(CurrentWord).Good = LearnDatabase(CurrentWord).Good + 1
            LearnDatabase(CurrentWord).GoodTotal = LearnDatabase(CurrentWord).GoodTotal + 1
            txtUser_Answer.BackColor = &HFF00&
            
            If GlobalGoOn Then
                If Len(Trim(txtUser_Answer.Text)) > 0 Or txtUser_Answer.BackColor = &HFF00& Then
                    CurrentWord = GetRandomWord
                End If
                ShowCurrentWord
            End If
        Else
            LearnDatabase(CurrentWord).Done = IIf(GlobalRepeat, False, True)
            LearnDatabase(CurrentWord).Wrong = LearnDatabase(CurrentWord).Wrong + 1
            LearnDatabase(CurrentWord).WrongTotal = LearnDatabase(CurrentWord).WrongTotal + 1
            txtUser_Answer.BackColor = &HFF&
        End If
    ElseIf KeyAscii = vbKeySpace And txtUser_Answer.BackColor <> &HFFFFFF Then
        If Len(Trim(txtUser_Answer.Text)) > 0 Or txtUser_Answer.BackColor = &HFF00& Then
            CurrentWord = GetRandomWord
        End If
        ShowCurrentWord
    End If
End Sub

Private Sub txtUser_Answer_KeyDown(KeyCode As Integer, Shift As Integer)
    RTFUser_Answer.RTFKeyDown KeyCode, Shift
End Sub

Private Sub txtUser_Answer_KeyUp(KeyCode As Integer, Shift As Integer)
    RTFUser_Answer.RTFKeyUp KeyCode, Shift
End Sub

Function GetQuestion(Word As Integer) As String
    GetQuestion = LearnDatabase(Word).Question
End Function

Function GetAnswer(Word As Integer) As String
    GetAnswer = LearnDatabase(Word).Answer
End Function

Function GetPronouncement(Word As Integer) As String
    GetPronouncement = LearnDatabase(Word).Pronouncement
End Function

Function GetLanguage(Word As Integer, Question As Boolean) As String
    GetLanguage = ""
    If Len(LearnDatabase(Word).Language) >= 2 Then
        If Question Then
            GetLanguage = Mid$(LearnDatabase(Word).Language, 1, 1)
        Else
            GetLanguage = Mid$(LearnDatabase(Word).Language, 2, 1)
        End If
    End If
End Function

Public Sub SetCurrentLearnMethod(myCurrentLearnMethod As Integer)
    CurrentLearnMethod = myCurrentLearnMethod
End Sub

Public Sub StartAgain()
    Dim i As Integer
    
    For i = 0 To LearnDatabaseCount - 1
            LearnDatabase(i).Good = 0
            LearnDatabase(i).Wrong = 0
            LearnDatabase(i).Done = False
    Next i
    
    CurrentWord = GetRandomWord
    ShowCurrentWord
End Sub

