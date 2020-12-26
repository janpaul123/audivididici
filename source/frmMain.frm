VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audivididici"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraAdvanced 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H000080FF&
      Height          =   4815
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exporteer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3000
         Width           =   2415
      End
      Begin VB.CommandButton cmdConvert 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Converteer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1920
         Width           =   2895
      End
      Begin VB.HScrollBar hsCompressionSlider 
         Height          =   255
         LargeChange     =   2
         Left            =   360
         Max             =   9
         TabIndex        =   26
         Top             =   2040
         Value           =   9
         Width           =   2775
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Terug"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label lblLink 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "www.audivididici.nl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4440
         TabIndex        =   37
         Top             =   3960
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":08CA
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
         Height          =   1095
         Left            =   360
         TabIndex        =   35
         Top             =   2880
         Width           =   4335
      End
      Begin VB.Label lblConvert 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Converteer bestanden van andere overhoorprogramma's (.oh4, .ohw,. t2k) naar Audivididici (.avd) bestanden."
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
         Height          =   975
         Left            =   4320
         TabIndex        =   31
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dit is Audivididici versie 0.00"
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
         Left            =   2760
         TabIndex        =   30
         Top             =   3720
         Width           =   4695
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "(C) Copyright by Simon Roosjen and Jan Paul Posma."
         ForeColor       =   &H0000AFF2&
         Height          =   255
         Left            =   2760
         TabIndex        =   29
         Top             =   4440
         Width           =   4695
      End
      Begin VB.Label lblCompressionSlider 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label lblCompression 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0959
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
         Height          =   1335
         Left            =   360
         TabIndex        =   27
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblAdvanced 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Geavanceerde opties"
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
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   2895
      End
      Begin VB.Shape shpAdvanced 
         BorderColor     =   &H000080FF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   4575
         Left            =   120
         Top             =   120
         Width           =   7455
      End
   End
   Begin VB.CheckBox chkSwitch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Wissel vraag en antwoord om"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   38
      Top             =   2520
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkGoOn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Direct doorgaan als het antwoord goed is"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   34
      Top             =   2280
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkRepeat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Herhaal totdat alles goed is"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   33
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkRandom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vragen husselen"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   39
      Top             =   1800
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CommandButton cmdAdvanced 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Geavanceerde opties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton optLearningMethod 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Uitspraak"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   21
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Start!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4560
      Width           =   3375
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Bewerk bestand"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Nieuw bestand"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bestand &openen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtFile 
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   5295
   End
   Begin VB.OptionButton optLearningMethod 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Diapresentatie"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   1560
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton optLearningMethod 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Oefenen"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1815
      Width           =   2175
   End
   Begin VB.OptionButton optLearningMethod 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Overschrijven"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   2085
      Width           =   2175
   End
   Begin VB.OptionButton optLearningMethod 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dictee"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   2355
      Width           =   2175
   End
   Begin VB.HScrollBar hsSpeedSlider 
      Height          =   255
      LargeChange     =   2
      Left            =   360
      Max             =   16
      Min             =   2
      TabIndex        =   5
      Top             =   3840
      Value           =   16
      Width           =   2775
   End
   Begin VB.CheckBox chkImages 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Met afbeeldingen"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chkSound 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Met geluid"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   2880
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblHelp 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Klik op 'Nieuw bestand' voor een nieuwe woordenlijst, of als je al een hebt op 'Bestand openen'."
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
      Height          =   735
      Left            =   4080
      TabIndex        =   12
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Shape shpHelp 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   3840
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label lblSettings 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Instellingen"
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
      Left            =   3960
      TabIndex        =   11
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblLearningMethod 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Methode"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblSpeed 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Snelheid"
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
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lblSpeedSlider 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Handmatig"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Shape shpMethod 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   120
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label lblMethodBack 
      BackStyle       =   0  'Transparent
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Shape shpSpeed 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Shape shpSettings 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   3840
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lblHelpBack 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   3840
      TabIndex        =   19
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label lblSettingsBack 
      BackStyle       =   0  'Transparent
      Height          =   1815
      Left            =   3840
      TabIndex        =   18
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lblSpeedBack 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const MAX_PATH As Long = 260

Private Declare Function SHGetPathFromIDList Lib "shell32" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" _
   Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" _
   (ByVal pv As Long)


Enum HighlightConstants
    hlNone = 0
    hlMethods = 1
    hlSpeed = 2
    hlSettings = 3
    hlHelp = 4
    hlOpenFile = 5
    hlNewFile = 6
    hlEditFile = 7
    hlStart = 8
    hlAdvanced = 9
    hlConvert = 10
    hlBack = 11
    hlExport = 12
End Enum

Private hlHighlight As HighlightConstants

Public DemoFile As String

Private Sub chkGoOn_Click()
    lblHelp.Caption = "Ga direct door bij een goed antwoord, of wacht tot je op de spatiebalk drukt."
End Sub

Private Sub chkImages_GotFocus()
    lblHelp.Caption = "Schakel afbeeldingen in of uit."
End Sub

Private Sub chkRandom_Click()
    lblHelp.Caption = "Hussel de vragen door elkaar zodat ze in willekeurige volgorde worden gevraagd."
End Sub

Private Sub chkRepeat_GotFocus()
    lblHelp.Caption = "Herhaal vragen tot je ze goed beantwoordt."
End Sub

Private Sub chkSound_GotFocus()
    lblHelp.Caption = "Schakel uitspraak en geluiden in of uit."
End Sub

Private Sub chkSwitch_Click()
    lblHelp.Caption = "Wissel de vraag en het antwoord om, om bij talen woordjes de andere kant op te leren."
End Sub

Private Sub cmdAdvanced_Click()
    fraAdvanced.Visible = True
    cmdBack.Default = True
    cmdBack.SetFocus
End Sub

Private Sub cmdAdvanced_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlAdvanced
End Sub

Private Sub cmdBack_Click()
    fraAdvanced.Visible = False
    NewFileSelected
End Sub

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlBack
End Sub

Private Sub cmdConvert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlConvert
End Sub

Private Sub cmdEdit_Click()
    On Error Resume Next
    
    Dim LastSlash As String
    Dim MyFile As String
    
    SetHighlight hlNone
    cmdStart.SetFocus
    
    If Demo Then
        MyFile = DemoFile
    Else
        MyFile = txtFile.Text
    End If
    
    If Len(MyFile) > 4 Then
        LastSlash = Len(MyFile) - InStr(1, StrReverse(MyFile), "\") + 1
        
        LoadWords Left$(MyFile, LastSlash), Right$(MyFile, Len(MyFile) - LastSlash), Left$(Right$(MyFile, Len(MyFile) - LastSlash), Len(MyFile) - LastSlash - 4), True
    End If
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlEditFile
End Sub

Private Sub cmdExport_Click()
    On Error Resume Next
    
    Dim Pictures As Boolean
    
    If Len(txtFile.Text) > 4 Then
        cdl.FileName = txtFile.Text
    End If
    
    cdl.DialogTitle = "Woordenlijst exporteren - Audivididici"
    cdl.FilterIndex = 0
    cdl.Filter = "Audivididici-woordenlijst (*.avd)|*.avd;"
    cdl.DefaultExt = ".avd"
    cdl.Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
    cdl.CancelError = True
    cdl.ShowOpen
    
    If Err.Number = cdlCancel Then
        Err.Clear
        Exit Sub
    End If
    
    If LCase$(Right$(cdl.FileName, 4)) <> ".avd" Or Not FExists(cdl.FileName) Then
        MsgBox "Selecteer een woordenlijst.", vbCritical + vbOKCancel, "Audivididici"
        Exit Sub
    Else
        Pictures = IIf(MsgBox("Wil je ook afbeeldingen exporteren?", vbQuestion + vbYesNo, "Audivididici") = vbYes, True, False)
        Load frmExport
        frmExport.Show
        frmExport.LoadFile Left$(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle)), cdl.FileTitle, Pictures
    End If
End Sub

Private Sub cmdExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlExport
End Sub

Private Sub cmdNew_Click()
    Load frmEdit
    frmEdit.New_File
    
    SetHighlight hlNone
End Sub

Private Sub cmdNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlNewFile
End Sub

Private Sub cmdOpen_Click()
    On Error Resume Next
    
    Dim LastSlash As Integer
    
    If Demo Then
        frmExample.Show vbModal
    Else
        If Len(txtFile.Text) > 4 Then
            cdl.FileName = txtFile.Text
        End If
        
        cdl.DialogTitle = "Woordenlijst openen - Audivididici"
        cdl.FilterIndex = 0
        cdl.Filter = "Audivididici-woordenlijsten (*.avd)|*.avd;|Teach2000 (*.t2k)|*.t2k;|Overhoor (*.oh4, *.ohw)|*.oh4;*.ohw;"
        cdl.DefaultExt = ".avd"
        cdl.Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
        cdl.CancelError = True
        cdl.ShowOpen
        
        If Err.Number = cdlCancel Then
            Err.Clear
            Exit Sub
        End If
        
        If LCase$(Right$(cdl.FileName, 4)) <> ".avd" And cdl.FileName <> "" Then
            If MsgBox("Wil je deze woordenlijst omzetten in een Audivididici-bestand?", vbQuestion + vbYesNo, "Audivididici") = vbYes Then
                txtFile.Text = Conversion.ConvertImport(Left$(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle)), cdl.FileTitle, hsCompressionSlider.value)
            Else
                txtFile.Text = ""
            End If
        Else
            txtFile.Text = cdl.FileName
        End If
    End If
    
    NewFileSelected
    
    SetHighlight hlNone
End Sub

Public Sub NewFileSelected()
    Dim MyFile As String
    On Error Resume Next
    
    If Demo Then
        MyFile = DemoFile
    Else
        MyFile = txtFile.Text
    End If
    
    If FExists(MyFile) Then
        lblHelp.Caption = "Klik hieronder op Start om te beginnen!"
        cmdEdit.Enabled = True
        cmdStart.Enabled = True
        cmdStart.Default = True
        cmdStart.SetFocus
    Else
        txtFile.Text = ""
        lblHelp.Caption = "Klik op 'Nieuw bestand' voor een nieuwe woordenlijst, of als je al een hebt op 'Bestand openen'."
        cmdEdit.Enabled = False
        cmdStart.Enabled = False
    End If
End Sub

Private Sub cmdOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlOpenFile
End Sub

Private Sub cmdStart_Click()
    Dim i As Integer
    Dim LastSlash As Integer
    Dim MyFile As String
    
    If Demo Then
        MyFile = DemoFile
    Else
        MyFile = txtFile.Text
    End If
    
    If Len(MyFile) > 4 Then
        For i = 0 To 4
            If optLearningMethod(i).value Then
                frmWords.SetCurrentLearnMethod (i)
            End If
        Next i
        
        LastSlash = Len(MyFile) - InStr(1, StrReverse(MyFile), "\") + 1
        
        LoadWords Left$(MyFile, LastSlash), Right$(MyFile, Len(MyFile) - LastSlash), Left$(Right$(MyFile, Len(MyFile) - LastSlash), Len(MyFile) - LastSlash - 4), False
    End If
    
    SetHighlight hlNone
End Sub

Private Sub cmdStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlStart
End Sub

Private Sub Form_Load()
    Me.Height = 5430
    fraAdvanced.Top = 120
    lblVersion.Caption = "Dit is " & App.ProductName & " versie " & App.Major & "." & App.Minor & App.Revision & "."
    If Demo Then
        frmNotification.Show vbModal
        cmdAdvanced.Enabled = False
    End If
End Sub

Private Sub fraAdvanced_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlNone
End Sub

Private Sub hsCompressionSlider_Change()
    lblCompressionSlider.Caption = Trim$(hsCompressionSlider.value)
End Sub

Private Sub hsSpeedSlider_GotFocus()
    lblHelp.Caption = "De computer kan automatisch doorklikken na een aantal seconden."
End Sub

Private Sub lblHelpBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlHelp
End Sub

Private Sub lblLink_Click()
    ShellExecute 0&, vbNullString, "http://www.audivididici.nl", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub lblMethodBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlMethods
End Sub

Private Sub SetHighlight(myHighlight As HighlightConstants)
    Dim i As Integer
    
    If myHighlight = hlHighlight Then Exit Sub
    
    Select Case hlHighlight
        Case hlMethods
            shpMethod.FillColor = &HFFFFFF
            For i = 0 To 4
                optLearningMethod(i).BackColor = &HFFFFFF
            Next i
        Case hlSpeed
            shpSpeed.FillColor = &HFFFFFF
        Case hlSettings
            shpSettings.FillColor = &HFFFFFF
            chkImages.BackColor = &HFFFFFF
            chkSound.BackColor = &HFFFFFF
            chkRepeat.BackColor = &HFFFFFF
            chkRandom.BackColor = &HFFFFFF
            chkGoOn.BackColor = &HFFFFFF
            chkSwitch.BackColor = &HFFFFFF
        Case hlHelp
            shpHelp.FillColor = &HFFFFFF
        Case hlOpenFile
            cmdOpen.BackColor = &HFFFFFF
        Case hlNewFile
            cmdNew.BackColor = &HFFFFFF
        Case hlEditFile
            cmdEdit.BackColor = &HFFFFFF
        Case hlStart
            cmdStart.BackColor = &HFFFFFF
        Case hlAdvanced
            cmdAdvanced.BackColor = &HFFFFFF
        Case hlConvert
            cmdConvert.BackColor = &HFFFFFF
        Case hlBack
            cmdBack.BackColor = &HFFFFFF
        Case hlExport
            cmdExport.BackColor = &HFFFFFF
    End Select
    
    Select Case myHighlight
        Case hlMethods
            shpMethod.FillColor = &HC0FFFF
            For i = 0 To 4
                optLearningMethod(i).BackColor = &HC0FFFF
            Next i
        Case hlSpeed
            shpSpeed.FillColor = &HC0FFFF
        Case hlSettings
            shpSettings.FillColor = &HC0FFFF
            chkImages.BackColor = &HC0FFFF
            chkSound.BackColor = &HC0FFFF
            chkRepeat.BackColor = &HC0FFFF
            chkRandom.BackColor = &HC0FFFF
            chkGoOn.BackColor = &HC0FFFF
            chkSwitch.BackColor = &HC0FFFF
        Case hlHelp
            shpHelp.FillColor = &HC0FFFF
        Case hlOpenFile
            cmdOpen.BackColor = &HC0FFFF
        Case hlNewFile
            cmdNew.BackColor = &HC0FFFF
        Case hlEditFile
            cmdEdit.BackColor = &HC0FFFF
        Case hlStart
            cmdStart.BackColor = &HC0FFFF
        Case hlAdvanced
            cmdAdvanced.BackColor = &HC0FFFF
        Case hlConvert
            cmdConvert.BackColor = &HC0FFFF
        Case hlBack
            cmdBack.BackColor = &HC0FFFF
        Case hlExport
            cmdExport.BackColor = &HC0FFFF
    End Select
    
    hlHighlight = myHighlight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmEdit
    Unload frmNotification
    Unload frmWords
    Unload frmResults
    Unload frmSplash
    Unload frmExport
    If Not Demo Then
        SaveSetting App.Title, "File", "txtFile.Text", txtFile.Text
        SaveSetting App.Title, "Compression", "hsCompressionSlider.Value", hsCompressionSlider.value
    End If
    SaveSetting App.Title, "LearningMethod", "optLearningMethod(0).Value", optLearningMethod(0).value
    SaveSetting App.Title, "LearningMethod", "optLearningMethod(1).Value", optLearningMethod(1).value
    SaveSetting App.Title, "LearningMethod", "optLearningMethod(2).Value", optLearningMethod(2).value
    SaveSetting App.Title, "LearningMethod", "optLearningMethod(3).Value", optLearningMethod(3).value
    SaveSetting App.Title, "LearningMethod", "optLearningMethod(4).Value", optLearningMethod(4).value
    SaveSetting App.Title, "Speed", "hsSpeedSlider.Value", hsSpeedSlider.value
    SaveSetting App.Title, "Settings", "chkImages.Value", chkImages.value
    SaveSetting App.Title, "Settings", "chkSound.Value", chkSound.value
    SaveSetting App.Title, "Settings", "chkRepeat.Value", chkRepeat.value
    SaveSetting App.Title, "Settings", "chkRandom.Value", chkRandom.value
    SaveSetting App.Title, "Settings", "chkGoOn.Value", chkGoOn.value
    SaveSetting App.Title, "Settings", "chkSwitch.Value", chkSwitch.value
    DelDir AppPath & "AVDTemp\"
    Dir AppPath
    End
End Sub

Private Sub hsSpeedSlider_Change()
    If hsSpeedSlider.value < 16 And hsSpeedSlider.Enabled = True Then
        lblSpeedSlider.Caption = Trim$(hsSpeedSlider.value)
    Else
        lblSpeedSlider.Caption = "Handmatig"
    End If
End Sub

Private Sub lblSettingsBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlSettings
End Sub

Private Sub lblSpeedBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlSpeed
End Sub

Function LoadWords(ByVal FilePath As String, FileName As String, FileDescription As String, Edit As Boolean) As Boolean
    If Right$(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    If Not FExists(FilePath & FileName) Then
        LoadWords = False
        Exit Function
    Else
        If Edit Then
            Load frmEdit
            If Not frmEdit.LoadFile(FilePath, FileName) Then
                Unload frmEdit
                LoadWords = False
                Exit Function
            Else
                cmdStart.Enabled = False
                cmdEdit.Enabled = False
                cmdAdvanced.Enabled = False
                frmEdit.Show
                Exit Function
            End If
        Else
            If Not frmWords.LoadFile(FilePath, FileName, FileDescription, IIf(hsSpeedSlider.Enabled, hsSpeedSlider.value, 16), chkImages Or optLearningMethod(0).value, chkSound Or optLearningMethod(4).value, chkRepeat, chkGoOn, chkSwitch And (optLearningMethod(1).value Or optLearningMethod(2).value Or optLearningMethod(3).value), chkRandom) Then
                frmWords.Hide
                LoadWords = False
                Exit Function
            Else
                cmdStart.Enabled = False
                cmdEdit.Enabled = False
                cmdAdvanced.Enabled = False
                frmWords.Show
                If optLearningMethod(2).value Or optLearningMethod(3).value Then frmWords.txtUser_Answer.SetFocus
                
                Exit Function
            End If
        End If
    End If
    LoadWords = True
End Function

Sub InitControls(myCommands As String)
    Dim i As Integer
    
    If Not Demo Then
        If Not FExists(myCommands) Then
            txtFile.Text = GetSetting(App.Title, "File", "txtFile.Text", "")
        Else
            txtFile.Text = myCommands
        End If
        
        hsCompressionSlider.value = Val(GetSetting(App.Title, "Compression", "hsCompressionSlider.Value", "5"))
    End If

    optLearningMethod(0).value = CBool(GetSetting(App.Title, "LearningMethod", "optLearningMethod(0).Value", "True"))
    optLearningMethod(1).value = CBool(GetSetting(App.Title, "LearningMethod", "optLearningMethod(1).Value", "False"))
    optLearningMethod(2).value = CBool(GetSetting(App.Title, "LearningMethod", "optLearningMethod(2).Value", "False"))
    optLearningMethod(3).value = CBool(GetSetting(App.Title, "LearningMethod", "optLearningMethod(3).Value", "False"))
    optLearningMethod(4).value = CBool(GetSetting(App.Title, "LearningMethod", "optLearningMethod(4).Value", "False"))
    
    hsSpeedSlider.value = Val(GetSetting(App.Title, "Speed", "hsSpeedSlider.Value", "16"))
    chkImages.value = Val(GetSetting(App.Title, "Settings", "chkImages.Value", "1"))
    chkSound.value = Val(GetSetting(App.Title, "Settings", "chkSound.Value", "1"))
    chkRepeat.value = Val(GetSetting(App.Title, "Settings", "chkRepeat.Value", "1"))
    chkRandom.value = Val(GetSetting(App.Title, "Settings", "chkRandom.Value", "1"))
    chkGoOn.value = Val(GetSetting(App.Title, "Settings", "chkGoOn.Value", "0"))
    chkSwitch.value = Val(GetSetting(App.Title, "Settings", "chkSwitch.Value", "0"))
    
    hsSpeedSlider_Change
    hsCompressionSlider_Change
    NewFileSelected
    
    For i = 0 To 4
        If optLearningMethod(i).value Then
            optLearningMethod_Click (i)
        End If
    Next i
End Sub

Private Sub optLearningMethod_Click(Index As Integer)
    Select Case Index
        Case 0
            hsSpeedSlider.Enabled = True
            chkImages.Enabled = False
            chkSound.Enabled = True
            chkRepeat.Enabled = False
            chkGoOn.Enabled = False
            chkSwitch.Enabled = False
            chkRandom.Enabled = True
        Case 1
            hsSpeedSlider.Enabled = True
            chkImages.Enabled = True
            chkSound.Enabled = True
            chkRepeat.Enabled = True
            chkGoOn.Enabled = False
            chkSwitch.Enabled = True
            chkRandom.Enabled = True
        Case 2
            hsSpeedSlider.Enabled = False
            chkImages.Enabled = True
            chkSound.Enabled = True
            chkRepeat.Enabled = True
            chkGoOn.Enabled = True
            chkSwitch.Enabled = True
            chkRandom.Enabled = True
        Case 3
            hsSpeedSlider.Enabled = False
            chkImages.Enabled = True
            chkSound.Enabled = True
            chkRepeat.Enabled = True
            chkGoOn.Enabled = True
            chkSwitch.Enabled = True
            chkRandom.Enabled = True
        Case 4
            hsSpeedSlider.Enabled = True
            chkImages.Enabled = True
            chkSound.Enabled = False
            chkRepeat.Enabled = True
            chkGoOn.Enabled = False
            chkSwitch.Enabled = False
            chkRandom.Enabled = True
    End Select
    hsSpeedSlider_Change
End Sub

Private Sub optLearningMethod_GotFocus(Index As Integer)
Select Case Index
        Case 0
            lblHelp.Caption = "Bij Diapresentatie kun je alle vragen bij langs lopen om ze te bekijken."
        Case 1
            lblHelp.Caption = "Bij Oefenen moet je het antwoord in je hoofd opzeggen."
        Case 2
            lblHelp.Caption = "Bij Overschrijven moet je het antwoord overtypen."
        Case 3
            lblHelp.Caption = "Bij Dictee moet je het antwoord foutloos intypen."
        Case 4
            lblHelp.Caption = "Bij Uitspraak moet je de vraag hardop uitspreken."
    End Select
End Sub

Private Sub txtFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight hlNone
End Sub

Private Sub cmdConvert_Click()
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim Path As String
    Dim pos As Long
    
    If MsgBox("Je kunt een map kiezen, waarna alle herkende bestanden worden omgezet in Audivididici-bestanden. Als je al Audivididici-bestanden in de geselecteerde map hebt staan, kunnen deze overschreven worden! Wil je doorgaan?", vbQuestion + vbYesNo, "Audivididici") = vbNo Then
        Exit Sub
    End If

    'Fill the BROWSEINFO structure with the
    'needed data. To accommodate comments, the
    'With/End With syntax has not been used, though
    'it should be your 'final' version.
    
    With bi
        'hwnd of the window that receives messages
        'from the call. Can be your application
        'or the handle from GetDesktopWindow()
        .hOwner = Me.hwnd
        
        'pointer to the item identifier list specifying
        'the location of the "root" folder to browse from.
        'If NULL, the desktop folder is used.
        .pidlRoot = 0&
        
        'message to be displayed in the Browse dialog
        .lpszTitle = "Selecteer de map met bestanden die je wil omzetten."
        
        'the type of folder to return.
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    'show the Browse Dialog
    pidl = SHBrowseForFolder(bi)
    
    'the dialog has closed, so parse & display the
    'user's returned folder selection contained in pidl
    Path = Space$(MAX_PATH)
    
    If SHGetPathFromIDList(ByVal pidl, ByVal Path) Then
        pos = InStr(Path, Chr$(0))
        Path = Left(Path, pos - 1)
        If Right(Path, 1) <> "\" Then Path = Path & "\"
        
        ConvertDir Path, hsCompressionSlider.value
    End If
    
    Call CoTaskMemFree(pidl)
End Sub
