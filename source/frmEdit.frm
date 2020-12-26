VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audivididici Creator"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10815
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDownload 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   2640
      TabIndex        =   51
      Top             =   3720
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label lblAdvanced 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Downloaden..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   0
         TabIndex        =   55
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "(C) Copyright by Simon Roosjen and Jan Paul Posma."
         ForeColor       =   &H0000AFF2&
         Height          =   255
         Left            =   2760
         TabIndex        =   54
         Top             =   4440
         Width           =   4695
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
         TabIndex        =   53
         Top             =   3720
         Width           =   4695
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
         TabIndex        =   52
         Top             =   3960
         Width           =   3015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   0
         Top             =   0
         Width           =   2775
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   8295
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9984
            MinWidth        =   3528
            Text            =   "Bestand:"
            TextSave        =   "Bestand:"
            Object.ToolTipText     =   "Bestand:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2461
            MinWidth        =   2469
            Text            =   "Totaal:"
            TextSave        =   "Totaal:"
            Object.ToolTipText     =   "Totaal aantal woorden:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "|||||||||||||||||||||||||||||||||||"
            TextSave        =   "|||||||||||||||||||||||||||||||||||"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "19:07"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDown 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Om&laag"
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
      Left            =   9480
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdUp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Om&hoog"
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
      Left            =   8040
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CheckBox chkShowExample 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Voorbeeld &tonen"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   48
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdSound_Download 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Download"
      Height          =   285
      Left            =   6960
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4600
      Width           =   855
   End
   Begin VB.CommandButton cmdImage_Download 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Download"
      Height          =   285
      Left            =   3000
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4600
      Width           =   855
   End
   Begin Audivididici.ctlClipboard Clip 
      Left            =   5760
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdSaveFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "O&pslaan"
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
      Left            =   6480
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5340
      Width           =   1455
   End
   Begin Audivididici.RTFUniversalUnicode RTFU_Answer 
      Height          =   2055
      Left            =   6720
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1695
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3625
   End
   Begin Audivididici.RTFUniversalUnicode RTFU_Question 
      Height          =   2055
      Left            =   6720
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   10
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3625
   End
   Begin RichTextLib.RichTextBox txtRaw_Answer 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1085
      _Version        =   393217
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   80
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmEdit.frx":08CA
   End
   Begin VB.CommandButton cmdDeleteWord 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Verwijder woord"
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
      Left            =   1920
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5340
      Width           =   1935
   End
   Begin VB.CommandButton cmdNewWord 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Nieuw woord"
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
      TabIndex        =   2
      Top             =   5340
      Width           =   1695
   End
   Begin VB.CommandButton cmdUndo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ongedaan maken"
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
      TabIndex        =   4
      Top             =   5340
      Width           =   2055
   End
   Begin VB.CommandButton cmdSound_Browse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bladeren"
      Height          =   285
      Left            =   6960
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4840
      Width           =   855
   End
   Begin VB.CommandButton cmdImage_Browse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bladeren"
      Height          =   285
      Left            =   3000
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4840
      Width           =   855
   End
   Begin VB.TextBox txtRaw_Sound 
      Height          =   285
      Left            =   4920
      TabIndex        =   23
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox txtRaw_Sound_Path 
      Height          =   285
      Left            =   4920
      TabIndex        =   24
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox txtRaw_Image 
      Height          =   285
      Left            =   1200
      TabIndex        =   20
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox txtRaw_Image_Path 
      Height          =   285
      Left            =   1200
      TabIndex        =   21
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CheckBox chkAnswerBold 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Vet"
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
      Left            =   240
      TabIndex        =   15
      Top             =   3030
      Width           =   615
   End
   Begin VB.CheckBox chkAnswerItalic 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cursief"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   3030
      Width           =   975
   End
   Begin VB.CheckBox chkAnswerUnderline 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Onderstrepen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   3030
      Width           =   1335
   End
   Begin VB.CheckBox chkAnswerStrikeThru 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doorhalen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   3030
      Width           =   1095
   End
   Begin VB.ComboBox cmbAnswer 
      Height          =   315
      ItemData        =   "frmEdit.frx":094C
      Left            =   5640
      List            =   "frmEdit.frx":0956
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CheckBox chkQuestionStrikeThru 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doorhalen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   1350
      Width           =   1095
   End
   Begin VB.CheckBox chkQuestionUnderline 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Onderstrepen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   1350
      Width           =   1335
   End
   Begin VB.CheckBox chkQuestionItalic 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cursief"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   1350
      Width           =   975
   End
   Begin VB.CheckBox chkQuestionBold 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Vet"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1350
      Width           =   735
   End
   Begin VB.ComboBox cmbQuestion 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmEdit.frx":0974
      Left            =   5640
      List            =   "frmEdit.frx":097E
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1320
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   8520
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrFillMainList 
      Interval        =   2000
      Left            =   8880
      Top             =   7320
   End
   Begin VB.Timer tmrSoundTracker 
      Interval        =   10
      Left            =   9240
      Top             =   7320
   End
   Begin VB.ListBox lstMainList 
      BackColor       =   &H00FFFFFF&
      Height          =   4545
      ItemData        =   "frmEdit.frx":099C
      Left            =   8040
      List            =   "frmEdit.frx":099E
      TabIndex        =   25
      Top             =   120
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox txtWord_Question 
      Height          =   975
      Left            =   2280
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6000
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1720
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"frmEdit.frx":09A0
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
   Begin RichTextLib.RichTextBox txtWord_Answer 
      Height          =   975
      Left            =   2280
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1720
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"frmEdit.frx":0A19
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
   Begin RichTextLib.RichTextBox tempRTF 
      Height          =   615
      Left            =   840
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmEdit.frx":0A92
   End
   Begin RichTextLib.RichTextBox txtRaw_Question 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1085
      _Version        =   393217
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   80
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmEdit.frx":0B14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Gentium"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MCI.MMControl MMCSound 
      Height          =   900
      Left            =   4200
      TabIndex        =   22
      Top             =   4080
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1588
      _Version        =   393216
      Orientation     =   1
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Image imgOK_Answer 
      Height          =   255
      Left            =   7440
      Picture         =   "frmEdit.frx":0B90
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Opslaan als:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   41
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Opzoeken:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   40
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Image imgOK_Sound 
      Height          =   255
      Left            =   7440
      Picture         =   "frmEdit.frx":0CD8
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgRaw_Image 
      Height          =   1095
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Opslaan als:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   39
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Opzoeken:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   38
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Image imgOK_Image 
      Height          =   255
      Left            =   3480
      Picture         =   "frmEdit.frx":0E20
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblAnswerFront 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "   "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   35
      Top             =   2925
      Width           =   255
   End
   Begin VB.Shape shpAnswerFront 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4920
      Top             =   2930
      Width           =   255
   End
   Begin VB.Label lblAnswerBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   36
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblQuestionFront 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "   "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   32
      Top             =   1245
      Width           =   255
   End
   Begin VB.Shape shpQuestionFront 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4920
      Top             =   1250
      Width           =   255
   End
   Begin VB.Label lblQuestionBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   33
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblAnswer 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Antwoord"
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
      TabIndex        =   37
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Image imgOK_Question 
      Height          =   255
      Left            =   7440
      Picture         =   "frmEdit.frx":0F68
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape shpAnswerBack 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5040
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vraag"
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
      TabIndex        =   34
      Top             =   240
      Width           =   2535
   End
   Begin VB.Shape shpQuestionBack 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5040
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image imgWord_Image 
      Height          =   2055
      Left            =   240
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Shape shpQuestion 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   7815
   End
   Begin VB.Shape shpAnswer 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   120
      Top             =   1800
      Width           =   7815
   End
   Begin VB.Label lblSound 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Uitspraak / Geluid"
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
      Left            =   4200
      TabIndex        =   43
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Shape shpSound 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   4080
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Shape shpExample 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2415
      Left            =   120
      Top             =   5880
      Width           =   10575
   End
   Begin VB.Label lblQuestionHighlight 
      BackStyle       =   0  'Transparent
      Height          =   1575
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label lblAnswerHighlight 
      BackStyle       =   0  'Transparent
      Height          =   1575
      Left            =   120
      TabIndex        =   45
      Top             =   1800
      Width           =   7815
   End
   Begin VB.Label lblSoundHighlight 
      BackStyle       =   0  'Transparent
      Height          =   1695
      Left            =   4080
      TabIndex        =   47
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label lblImage 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Afbeelding"
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
      TabIndex        =   42
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Shape shpImage 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   120
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label lblImageHighlight 
      BackStyle       =   0  'Transparent
      Height          =   1695
      Left            =   120
      TabIndex        =   46
      Top             =   3480
      Width           =   3855
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type LearnType
    Question As String
    Answer As String
    Image As String
    Sound As String
    Language As String
End Type

Private Enum EditHighlightConstants
    ehlNone = 0
    ehlQuestion = 1
    ehlAnswer = 2
    ehlImage = 3
    ehlSound = 4
    ehlUndo = 5
    ehlNewWord = 6
    ehlDeleteWord = 7
    ehlMainList = 8
    ehlsavefile = 9
End Enum

Private Enum EditHightlightTypes
    ehlFocus = 0
    ehlMouse = 1
End Enum

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" _
    (ByVal lpFileName As String) As Long
    
Private Declare Function URLDownloadToCacheFile Lib "urlmon" Alias "URLDownloadToCacheFileA" (ByVal lpUnkcaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwBufLength As Long, ByVal dwReserved As Long, ByVal IBindStatusCallback As Long) As Long
 
    
Const imgRaw_Image_Width = 855
Const imgRaw_Image_Height = 975
Const imgRaw_Image_Top = 3960
Const imgRaw_Image_Left = 240
Const imgWord_Image_Width = 1815
Const imgWord_Image_Height = 2055
Const imgWord_Image_Top = 6000
Const imgWord_Image_Left = 240

Private LearnDatabase(999) As LearnType
Private LearnDatabaseCount As Integer

Private CurrentWord As Integer
Private GlobalFilePath As String
Private GlobalFileName As String
Private Recording As String

Private txtRaw_Question_SelChanging As Boolean
Private txtRaw_Question_Changing As Boolean
Private txtRaw_Answer_SelChanging As Boolean
Private txtRaw_Answer_Changing As Boolean

Private FocusHighlight As EditHighlightConstants
Private MouseHighlight As EditHighlightConstants

Private FileChanged As Boolean
Private FileChangedBeforeWord As Boolean

Private Compression As Integer

Private OK_Question As Boolean
Private OK_Answer As Boolean
Private OK_Image As Boolean
Private OK_Sound As Boolean


Private Sub NewWordSelected()
    On Error Resume Next
    
    Dim TempSelStart, TempSelLength As Integer
    
    Recording = ""
    MMCSound.Command = "stop"
    MMCSound.Command = "close"
    
    DeleteFile AppPath & "TempRecord.wav"
    
    LearnDatabase(CurrentWord).Question = CleanRTF(LearnDatabase(CurrentWord).Question)
    LearnDatabase(CurrentWord).Answer = CleanRTF(LearnDatabase(CurrentWord).Answer)
    
    TempSelStart = txtRaw_Question.SelStart
    TempSelLength = txtRaw_Question.SelLength
    If Left$(LearnDatabase(CurrentWord).Question, 6) <> "{\rtf1" Then
        txtRaw_Question.TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}}\viewkind4\pard\b\f0\fs38\qc " & LearnDatabase(CurrentWord).Question & "\par}"
    Else
        txtRaw_Question.TextRTF = LearnDatabase(CurrentWord).Question
    End If
    
    txtRaw_Question.SelStart = 0
    txtRaw_Question.SelLength = Len(txtRaw_Question.Text)
    txtRaw_Question.SelFontName = "Gentium"
    txtRaw_Question.SelFontSize = 18
    txtRaw_Question.SelAlignment = rtfCenter
    
    txtRaw_Question.SelStart = TempSelStart
    txtRaw_Question.SelLength = TempSelLength
    
    TempSelStart = txtRaw_Answer.SelStart
    TempSelLength = txtRaw_Answer.SelLength
    If Left$(LearnDatabase(CurrentWord).Answer, 6) <> "{\rtf1" Then
        txtRaw_Answer.TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}}\viewkind4\uc1\pard\f0\fs32\qc " & LearnDatabase(CurrentWord).Answer & "\par}"
    Else
        txtRaw_Answer.TextRTF = LearnDatabase(CurrentWord).Answer
    End If
    
    txtRaw_Answer.SelStart = 0
    txtRaw_Answer.SelLength = Len(txtRaw_Answer.Text)
    txtRaw_Answer.SelFontName = "Gentium"
    txtRaw_Answer.SelFontSize = 18
    txtRaw_Answer.SelAlignment = rtfCenter
    
    txtRaw_Answer.SelStart = TempSelStart
    txtRaw_Answer.SelLength = TempSelLength
    
    If Len(LearnDatabase(CurrentWord).Language) >= 2 Then
        If Mid$(LearnDatabase(CurrentWord).Language, 1, 1) = "L" Then
            cmbQuestion.ListIndex = 0
        ElseIf Mid$(LearnDatabase(CurrentWord).Language, 1, 1) = "G" Then
            cmbQuestion.ListIndex = 1
        End If
    
        If Mid$(LearnDatabase(CurrentWord).Language, 2, 1) = "L" Then
            cmbAnswer.ListIndex = 0
        ElseIf Mid$(LearnDatabase(CurrentWord).Language, 2, 1) = "G" Then
            cmbAnswer.ListIndex = 1
        End If
    Else
        If cmbQuestion.ListIndex < 0 Or cmbQuestion.ListIndex > 1 Then
            cmbQuestion.ListIndex = 0
        End If
        If cmbAnswer.ListIndex < 0 Or cmbAnswer.ListIndex > 1 Then
            cmbAnswer.ListIndex = 0
        End If
    End If
    
    txtRaw_Question.Locked = False
    txtRaw_Answer.Locked = False
    
    txtRaw_Image.Text = LearnDatabase(CurrentWord).Image
    txtRaw_Sound.Text = LearnDatabase(CurrentWord).Sound
    
    txtRaw_Image_Path.Text = ""
    txtRaw_Sound_Path.Text = ""
    
    OK_Question = True
    OK_Answer = True
    
    SetRawSound
    SetRawImage
    
    FileChanged = FileChangedBeforeWord
    
    txtWord_Question.TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}}\viewkind4\pard\f0\fs38\qc " & CleanRTF(txtRaw_Question.TextRTF) & "\par}"
    txtWord_Answer.TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}}\viewkind4\uc1\pard\f0\fs32\qc " & CleanRTF(txtRaw_Answer.TextRTF) & "\par}"
    
    MakePath False
    
    Form_Refresh
End Sub

Private Sub Form_Refresh()
    If Not (Me.WindowState <> vbMinimized And Me.Visible = True) Then Exit Sub
    
    With Me
        StatusBar.Panels.Item(1).Text = "Bestand: " & GlobalFileName
        StatusBar.Panels.Item(1).ToolTipText = "Bestand: " & GlobalFileName
        StatusBar.Panels.Item(2).Text = "Totaal: " & LearnDatabaseCount & " "
        StatusBar.Panels.Item(2).ToolTipText = "Totaal aantal Worden: " & LearnDatabaseCount & " - Klik hier!"
        
        .Caption = GlobalFileName & IIf(FileChanged, " [Aangepast]", "") & " - Audivididici Creator"
    End With
End Sub

Private Sub chkAnswerBold_GotFocus()
    SetHighlight ehlAnswer, ehlFocus
End Sub

Private Sub chkAnswerItalic_Click()
    On Error Resume Next
    txtRaw_Answer_Changing = True
    If Not txtRaw_Answer_SelChanging Then
        txtRaw_Answer.SelItalic = IIf(chkAnswerItalic.value = 1, True, False)
        RTFU_Answer.RTFSetFont chkAnswerBold.value, chkAnswerItalic.value, chkAnswerUnderline.value, chkAnswerStrikeThru.value
        txtRaw_Answer.SetFocus
    End If
    txtRaw_Answer_Changing = False
End Sub

Private Sub chkAnswerItalic_GotFocus()
    SetHighlight ehlAnswer, ehlFocus
End Sub

Private Sub chkAnswerStrikeThru_Click()
    On Error Resume Next
    txtRaw_Answer_Changing = True
    If Not txtRaw_Answer_SelChanging Then
        txtRaw_Answer.SelStrikeThru = IIf(chkAnswerStrikeThru.value = 1, True, False)
        RTFU_Answer.RTFSetFont chkAnswerBold.value, chkAnswerItalic.value, chkAnswerUnderline.value, chkAnswerStrikeThru.value
        txtRaw_Answer.SetFocus
    End If
    txtRaw_Answer_Changing = False
End Sub

Private Sub chkAnswerStrikeThru_GotFocus()
    SetHighlight ehlAnswer, ehlFocus
End Sub

Private Sub chkAnswerUnderline_Click()
    On Error Resume Next
    txtRaw_Answer_Changing = True
    If Not txtRaw_Answer_SelChanging Then
        txtRaw_Answer.SelUnderline = IIf(chkAnswerUnderline.value = 1, True, False)
        RTFU_Answer.RTFSetFont chkAnswerBold.value, chkAnswerItalic.value, chkAnswerUnderline.value, chkAnswerStrikeThru.value
        txtRaw_Answer.SetFocus
    End If
    txtRaw_Answer_Changing = False
End Sub

Private Sub chkAnswerBold_Click()
    On Error Resume Next
    txtRaw_Answer_Changing = True
    If Not txtRaw_Answer_SelChanging Then
        txtRaw_Answer.SelBold = IIf(chkAnswerBold.value = 1, True, False)
        RTFU_Answer.RTFSetFont chkAnswerBold.value, chkAnswerItalic.value, chkAnswerUnderline.value, chkAnswerStrikeThru.value
        txtRaw_Answer.SetFocus
    End If
    txtRaw_Answer_Changing = False
End Sub


Private Sub chkAnswerUnderline_GotFocus()
    SetHighlight ehlAnswer, ehlFocus
End Sub

Private Sub chkQuestionBold_GotFocus()
    SetHighlight ehlQuestion, ehlFocus
End Sub

Private Sub chkQuestionItalic_Click()
    On Error Resume Next
    txtRaw_Question_Changing = True
    If Not txtRaw_Question_SelChanging Then
        txtRaw_Question.SelItalic = IIf(chkQuestionItalic.value = 1, True, False)
        RTFU_Question.RTFSetFont chkQuestionBold.value, chkQuestionItalic.value, chkQuestionUnderline.value, chkQuestionStrikeThru.value
        txtRaw_Question.SetFocus
    End If
    txtRaw_Question_Changing = False
End Sub

Private Sub chkQuestionItalic_GotFocus()
    SetHighlight ehlQuestion, ehlFocus
End Sub

Private Sub chkQuestionStrikeThru_Click()
    On Error Resume Next
    txtRaw_Question_Changing = True
    If Not txtRaw_Question_SelChanging Then
        txtRaw_Question.SelStrikeThru = IIf(chkQuestionStrikeThru.value = 1, True, False)
        RTFU_Question.RTFSetFont chkQuestionBold.value, chkQuestionItalic.value, chkQuestionUnderline.value, chkQuestionStrikeThru.value
        txtRaw_Question.SetFocus
    End If
    txtRaw_Question_Changing = False
End Sub

Private Sub chkQuestionStrikeThru_GotFocus()
    SetHighlight ehlQuestion, ehlFocus
End Sub

Private Sub chkQuestionUnderline_Click()
    On Error Resume Next
    txtRaw_Question_Changing = True
    If Not txtRaw_Question_SelChanging Then
        txtRaw_Question.SelUnderline = IIf(chkQuestionUnderline.value = 1, True, False)
        RTFU_Question.RTFSetFont chkQuestionBold.value, chkQuestionItalic.value, chkQuestionUnderline.value, chkQuestionStrikeThru.value
        txtRaw_Question.SetFocus
    End If
    txtRaw_Question_Changing = False
End Sub

Private Sub chkQuestionBold_Click()
    On Error Resume Next
    txtRaw_Question_Changing = True
    If Not txtRaw_Question_SelChanging Then
        txtRaw_Question.SelBold = IIf(chkQuestionBold.value = 1, True, False)
        RTFU_Question.RTFSetFont chkQuestionBold.value, chkQuestionItalic.value, chkQuestionUnderline.value, chkQuestionStrikeThru.value
        txtRaw_Question.SetFocus
    End If
    txtRaw_Question_Changing = False
End Sub

Public Sub New_File()
    MMCSound.Command = "close"
    DelDir AppPath & "AVDTemp\"
    Dir AppPath
    MkDir AppPath & "AVDTemp\"
    GlobalFilePath = ""
    GlobalFileName = "Nieuw Bestand"
    
    Me.Show
    
    ClearDatabase
    
    OK_Question = True
    OK_Answer = True
    
    cmdNewWord_Click
    
    FileChanged = False
End Sub

Function SetNewFile() As Boolean
    On Error Resume Next
    
    Dim FilePath, FileName As String
    
    cdl.DialogTitle = "Woordenlijst opslaan - Audivididici"
    cdl.FilterIndex = 0
    cdl.Filter = "Woordenlijsten (*.avd)|*.avd"
    cdl.DefaultExt = ".avd"
    cdl.Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
    cdl.CancelError = True
    cdl.ShowSave
    
    If Err.Number = cdlCancel Then
        Err.Clear
        SetNewFile = False
        Exit Function
    End If
    
    If Not (LCase$(Right$(cdl.FileName, 4)) = ".avd") Then
        SetNewFile = False
        Exit Function
    End If
    
    FilePath = Left$(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle))
    FileName = cdl.FileTitle
    
    If Right$(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    
    If FExists(FilePath & FileName) Then
        If MsgBox(FileName & " bestaat al. Wil je dit bestand overschrijven?", vbQuestion + vbYesNo, "Audivididici Creator") = vbNo Then
            SetNewFile = False
            Exit Function
        End If
    End If
    
    GlobalFilePath = FilePath
    GlobalFileName = FileName
    
    SetNewFile = True
End Function

Private Sub chkQuestionUnderline_GotFocus()
    SetHighlight ehlQuestion, ehlFocus
End Sub

Private Sub chkShowExample_Click()
    If chkShowExample Then
        Me.Height = 9000
    Else
        Me.Height = 6550
    End If
End Sub

Private Sub Clip_ClipboardChanged()
    Dim myText As String
    
    myText = Clipboard.GetText
    
    If Len(myText) > 4 Then
        If (Right$(myText, 4) = ".jpg" Or Right$(myText, 5) = ".jpeg" Or Right$(myText, 4) = ".bmp" Or Right$(myText, 4) = ".gif") Then
            If Len(txtRaw_Image_Path.Text) <= 0 Then
                txtRaw_Image_Path.Text = myText
            End If
        ElseIf (Right$(myText, 4) = ".wav" Or Right$(myText, 4) = ".mp3") Then
            If Len(txtRaw_Sound_Path.Text) <= 0 Then
                txtRaw_Sound_Path.Text = myText
            End If
        End If
    End If
End Sub

Private Sub cmbAnswer_Change()
    OK_Answer = False
    If Not FileChanged Then
        FileChanged = True
        Form_Refresh
    End If
End Sub

Private Sub cmbAnswer_Click()
    RTFU_Answer.SetKeyboard cmbAnswer.ListIndex
End Sub

Private Sub cmbAnswer_GotFocus()
    SetHighlight ehlAnswer, ehlFocus
End Sub

Private Sub cmbQuestion_Change()
    OK_Question = False
    If Not FileChanged Then
        FileChanged = True
        Form_Refresh
    End If
End Sub

Private Sub cmbQuestion_Click()
    RTFU_Question.SetKeyboard cmbQuestion.ListIndex
End Sub

Private Sub cmbQuestion_GotFocus()
    SetHighlight ehlQuestion, ehlFocus
End Sub

Private Sub cmdDeleteWord_GotFocus()
    SetHighlight ehlDeleteWord, ehlFocus
End Sub

Private Sub cmdDeleteWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight ehlDeleteWord, ehlMouse
End Sub

Private Sub cmdDown_Click()
    Dim TempLearnType As LearnType
    Dim TempList As String
    
    If CurrentWord >= LearnDatabaseCount - 1 Or LearnDatabaseCount <= 1 Then Exit Sub
    
    CheckSaveWord
    
    TempLearnType = LearnDatabase(CurrentWord)
    LearnDatabase(CurrentWord) = LearnDatabase(CurrentWord + 1)
    LearnDatabase(CurrentWord + 1) = TempLearnType
    
    TempList = lstMainList.List(CurrentWord)
    lstMainList.List(CurrentWord) = lstMainList.List(CurrentWord + 1)
    lstMainList.List(CurrentWord + 1) = TempList
    
    CurrentWord = CurrentWord + 1
    
    lstMainList.Selected(CurrentWord) = True
    NewWordSelected
    
    FileChanged = True
    Form_Refresh
End Sub

Private Sub cmdImage_Browse_Click()
    On Error Resume Next

    If LearnDatabaseCount <= 0 Then Exit Sub

    cdl.DialogTitle = "Afbeelding zoeken - Audivididici"
    cdl.FilterIndex = 0
    cdl.Filter = "Afbeeldingen (*.jpg;*.jpeg;*.gif;*.bmp)|*.jpg;*.jpeg;*.gif;*.bmp"
    cdl.DefaultExt = ".jpg"
    cdl.Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
    cdl.CancelError = True
    cdl.ShowOpen
    
    If Err.Number = cdlCancel Then
        Err.Clear
        Exit Sub
    End If
    
    If Not (LCase$(Right$(cdl.FileName, 4)) = ".jpg" Or LCase$(Right$(cdl.FileName, 5)) = ".jpeg" Or LCase$(Right$(cdl.FileName, 4)) = ".gif" Or LCase$(Right$(cdl.FileName, 4)) = ".bmp") Then
        MsgBox "Je hebt geen geldig bestand geselecteerd!", vbCritical + vbOKOnly, "Audivididici Creator"
        Exit Sub
    End If
    txtRaw_Image_Path.Text = cdl.FileName
    SetRawImage
End Sub

Private Sub cmdImage_Browse_GotFocus()
    SetHighlight ehlImage, ehlFocus
End Sub

Private Sub cmdImage_Download_Click()
    If Len(txtRaw_Image_Path.Text) > 0 And Len(txtRaw_Image.Text) > 0 And ValidPath(txtRaw_Image.Text) Then
        txtRaw_Image_Path.Text = DownloadFile(txtRaw_Image_Path.Text)
    ElseIf Len(txtRaw_Image_Path.Text) <= 0 Then
        MsgBox "Kopiëer een internetadres naar het vakje 'Opzoeken'. (Bijvoorbeeld: http://www.audivididici.nl/Audivididici.gif)", vbInformation + vbOKOnly, "Audivididici"
    End If
End Sub

Private Sub cmdNewWord_Click()
    If LearnDatabaseCount > 0 Then
        CheckSaveWord
    End If
    
    LearnDatabaseCount = LearnDatabaseCount + 1
    CurrentWord = LearnDatabaseCount - 1
    LearnDatabase(CurrentWord).Question = " "
    LearnDatabase(CurrentWord).Answer = "???"
    LearnDatabase(CurrentWord).Image = ""
    LearnDatabase(CurrentWord).Sound = ""
    LearnDatabase(CurrentWord).Question = ""
    
    lstMainList.AddItem ("???")
    lstMainList.Selected(lstMainList.ListCount - 1) = True
    
    NewWordSelected
    
    txtRaw_Question.SelStart = 0
    txtRaw_Question.SelLength = 1
    txtRaw_Question.SelBold = True
    txtRaw_Question.SelColor = &H80FF&
    
    txtRaw_Question.SetFocus
End Sub

Private Sub cmdNewWord_GotFocus()
    SetHighlight ehlNewWord, ehlFocus
End Sub

Private Sub cmdNewWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight ehlNewWord, ehlMouse
End Sub

Private Function SaveFile() As Boolean
    Dim i As Integer
    Dim MyFile As Integer
    Dim RetCode As Integer
    Dim MsgboxRet As VbMsgBoxResult
    Dim SetFileRet As Boolean
    
    If Not DirExists(GlobalFilePath) Then
        If Not SetNewFile Then
            SaveFile = False
            Exit Function
        End If
    End If
    
    If Demo Then
        Exit Function
    End If
    
    Load frmSplash
    frmSplash.SetMaximumAndValue LearnDatabaseCount, 0
    frmSplash.Show
    
    If FExists(AppPath & "AVDTemp\info.txt") Then
        DeleteFile AppPath & "AVDTemp\info.txt"
    End If
    
    MyFile = FreeFile
    
    Open AppPath & "AVDTemp\info.txt" For Output As MyFile
        Print #MyFile, Trim$(Str$(LearnDatabaseCount))
        
        For i = 0 To LearnDatabaseCount - 1
            Print #MyFile, Trim$(CleanRTF(LearnDatabase(i).Question))
            Print #MyFile, Trim$(CleanRTF(LearnDatabase(i).Answer))
            Print #MyFile, Trim$(CleanRTF(LearnDatabase(i).Image))
            Print #MyFile, Trim$(CleanRTF(LearnDatabase(i).Sound))
            Print #MyFile, Trim$(CleanRTF(LearnDatabase(i).Language))
            
            frmSplash.SetValue i + 1
            
            DoEvents
        Next i
    Close MyFile
        
    
    If FExists(GlobalFilePath & GlobalFileName) Then
        DeleteFile GlobalFilePath & GlobalFileName
    End If
    
    '-- Set Options - Only The Common Ones Are Shown Here
    '-- These Must Be Set Before Calling The VBZip32 Function
    zDate = vbNullString
    'zDate = "2005-1-31"
    'zExcludeDate = 1
    'zIncludeDate = 0
    zJunkDir = 1     ' 1 = Throw Away Path Names
    zRecurse = 0     ' 1 = Recurse -r 2 = Recurse -R 2 = Most Useful :)
    zUpdate = 0      ' 1 = Update Only If Newer
    zFreshen = 0     ' 1 = Freshen - Overwrite Only
    zLevel = Asc(Compression)  ' Compression Level (0 - 9)
    zEncrypt = 0     ' Encryption = 1 For Password Else 0
    zComment = 0     ' Comment = 1 if required
    
    '-- Select Some Files - Wildcards Are Supported
    '-- Change The Paths Here To Your Directory
    '-- And Files!!!
    ' Change ZIPnames in VBZipBas.bas if need more than 100 files
    zArgc = 1           ' Number Of Elements Of mynames Array
    zZipFileName = GlobalFilePath & GlobalFileName
    zZipFileNames.zFiles(0) = AppPath & "AVDTemp\*.*"
    zRootDir = AppPath & "AVDTemp\"    ' This Affects The Stored Path Name
    
    ' Older versions of Zip32.dll do not handle setting
    ' zRootDir to anything other than "".  If you need to
    ' change root directory an alternative is to just change
    ' directory.  This requires Zip32.dll to be on the command
    ' path.  This should be fixed in Zip 2.31.  1/31/2005 EG
    
    ' ChDir "a"
    
    '-- Go Zip Them Up!
    RetCode = VBZip32
    
    FillMainList
    FileChanged = False
    FileChangedBeforeWord = False
    Form_Refresh
    
    Unload frmSplash
    
    SaveFile = True
End Function

Private Sub cmdSaveFile_Click()
    CheckSaveWord
    SaveFile
End Sub

Private Sub cmdSaveFile_GotFocus()
    SetHighlight ehlsavefile, ehlFocus
End Sub

Private Sub cmdSaveFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight ehlsavefile, ehlMouse
End Sub

Private Sub SaveWord()
    On Error Resume Next
    Dim myQuestion As String
    Dim myAnswer As String
    
    LearnDatabase(CurrentWord).Question = CleanRTF(txtRaw_Question.TextRTF)
    LearnDatabase(CurrentWord).Answer = CleanRTF(txtRaw_Answer.TextRTF)
    
    If cmbQuestion.ListIndex = 0 Then
        LearnDatabase(CurrentWord).Language = "L"
    ElseIf cmbQuestion.ListIndex = 1 Then
        LearnDatabase(CurrentWord).Language = "G"
    Else
        LearnDatabase(CurrentWord).Language = " "
    End If
    
    If cmbAnswer.ListIndex = 0 Then
        LearnDatabase(CurrentWord).Language = LearnDatabase(CurrentWord).Language & "L"
    ElseIf cmbAnswer.ListIndex = 1 Then
        LearnDatabase(CurrentWord).Language = LearnDatabase(CurrentWord).Language & "G"
    Else
        LearnDatabase(CurrentWord).Language = LearnDatabase(CurrentWord).Language & " "
    End If
    
    If LearnDatabaseCount <= 0 Then Exit Sub
    
    If FExists(txtRaw_Image_Path.Text) And Len(txtRaw_Image.Text) > 0 And ValidPath(AppPath & "AVDTemp\" & txtRaw_Image.Text) Then
        FileCopy txtRaw_Image_Path.Text, AppPath & "AVDTemp\" & txtRaw_Image.Text
        LearnDatabase(CurrentWord).Image = txtRaw_Image.Text
        txtRaw_Image_Path.Text = ""
    ElseIf txtRaw_Image.Text <> LearnDatabase(CurrentWord).Image And FExists(AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Image) And ValidPath(AppPath & "AVDTemp\" & txtRaw_Image.Text) And Len(txtRaw_Image.Text) > 0 Then
        FileCopy AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Image, AppPath & "AVDTemp\" & txtRaw_Image.Text
        DeleteFile AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Image
        LearnDatabase(CurrentWord).Image = txtRaw_Image.Text
        txtRaw_Image_Path.Text = ""
    ElseIf txtRaw_Image.Text <> LearnDatabase(CurrentWord).Image And FExists(AppPath & "AVDTemp\" & txtRaw_Image.Text) Then
        LearnDatabase(CurrentWord).Image = txtRaw_Image.Text
        txtRaw_Image_Path.Text = ""
    End If

    If FExists(txtRaw_Sound_Path.Text) And Len(txtRaw_Sound.Text) > 0 And ValidPath(AppPath & "AVDTemp\" & txtRaw_Sound.Text) Then
        FileCopy txtRaw_Sound_Path.Text, AppPath & "AVDTemp\" & txtRaw_Sound.Text
        LearnDatabase(CurrentWord).Sound = txtRaw_Sound.Text
        txtRaw_Sound_Path.Text = ""
    ElseIf FExists(AppPath & "TempRecord.wav") And Len(txtRaw_Sound.Text) > 0 And ValidPath(AppPath & "AVDTemp\" & txtRaw_Sound.Text) Then
        MMCSound.Command = "stop"
        MMCSound.Command = "close"
    
        FileCopy AppPath & "TempRecord.wav", AppPath & "AVDTemp\" & txtRaw_Sound.Text
        DeleteFile AppPath & "TempRecord.wav"
        Recording = ""
        LearnDatabase(CurrentWord).Sound = txtRaw_Sound.Text
        SetRawSound
    ElseIf (txtRaw_Sound.Text <> LearnDatabase(CurrentWord).Sound) And FExists(AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Sound) And ValidPath(AppPath & "AVDTemp\" & txtRaw_Sound.Text) And Len(txtRaw_Sound.Text) > 0 Then
        FileCopy AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Sound, AppPath & "AVDTemp\" & txtRaw_Sound.Text
        DeleteFile AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Sound
        LearnDatabase(CurrentWord).Sound = txtRaw_Sound.Text
        txtRaw_Sound_Path.Text = ""
    ElseIf (txtRaw_Sound.Text <> LearnDatabase(CurrentWord).Sound) And FExists(AppPath & "AVDTemp\" & txtRaw_Sound.Text) Then
        txtRaw_Sound_Path.Text = ""
        LearnDatabase(CurrentWord).Sound = txtRaw_Sound.Text
    End If
    
    SetRawImage
    SetRawSound
    OK_Question = True
    OK_Answer = True
    
    myQuestion = Replace(NormalText(LearnDatabase(CurrentWord).Question), "?", "")
    myAnswer = Replace(NormalText(LearnDatabase(CurrentWord).Answer), "?", "")
                
    lstMainList.List(CurrentWord) = ListMarkup(CurrentWord + 1, myQuestion, myAnswer, 32)
    
    FileChanged = True
    FileChangedBeforeWord = True
    Form_Refresh
End Sub

Function ListMarkup(Number As Integer, Question As String, Answer As String, Length As Integer) As String
    Dim myQuestion As String
    Dim myAnswer As String
    Dim myLength As Integer
    Dim mySeperator As String
    
    myQuestion = LettersOnly(Question)
    myAnswer = LettersOnly(Answer)
    myLength = Length - 3
    
    If Len(myQuestion) > 0 And Len(myAnswer) > 0 Then
        mySeperator = " - "
        myLength = myLength - 3
    End If
    
    If Len(myQuestion) + Len(myAnswer) > myLength Then
        If Len(myQuestion) > myLength / 2 Then
            myQuestion = TrimNicely(myQuestion, myLength / 2)
        End If
        If Len(myAnswer) > myLength - Len(myQuestion) Then
            myAnswer = TrimNicely(myAnswer, myLength - Len(myQuestion))
        End If
    End If
    
    ListMarkup = Str$(Number) & ". " & myQuestion & mySeperator & myAnswer
End Function

Function LettersOnly(MyString As String) As String
    Dim OutputString As String
    Dim i As Integer
    Dim CurChar As String
    Dim HadLetter As Boolean
    
    HadLetter = False
    MyString = Trim(MyString)
    
    For i = 1 To Len(MyString)
        CurChar = Mid(MyString, i, 1)
        
        If Len(CurChar) > 0 Then
            If Asc(LCase(CurChar)) >= 97 And Asc(LCase(CurChar)) <= 122 Then
                HadLetter = True
                
                OutputString = OutputString & CurChar
            ElseIf Asc(LCase(CurChar)) >= 32 And Asc(LCase(CurChar)) <= 126 Then
                OutputString = OutputString & CurChar
            End If
        End If
    Next i
    
    If HadLetter Then
        LettersOnly = OutputString
    Else
        LettersOnly = ""
    End If
End Function

Private Sub cmdUndo_GotFocus()
    SetHighlight ehlUndo, ehlFocus
End Sub

Private Sub cmdUndo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight ehlUndo, ehlMouse
End Sub

Private Sub cmdSound_Browse_Click()
    On Error Resume Next
        
    If LearnDatabaseCount <= 0 Then Exit Sub

    cdl.DialogTitle = "Geluid zoeken - Audivididici"
    cdl.FilterIndex = 0
    cdl.Filter = "Geluid (*.wav;*.mp3)|*.wav;*.mp3"
    cdl.DefaultExt = ".wav"
    cdl.Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
    cdl.CancelError = True
    cdl.ShowOpen
    
    If Err.Number = cdlCancel Then
        Err.Clear
        Exit Sub
    End If
    
    If Not (LCase$(Right$(cdl.FileName, 4)) = ".wav" Or LCase$(Right$(cdl.FileName, 4)) = ".mp3") Then
        MsgBox "Je hebt geen geldig bestand geselecteerd!", vbCritical + vbOKOnly, "Audivididici Creator"
        Exit Sub
    End If
    txtRaw_Sound_Path.Text = cdl.FileName
    SetRawSound
End Sub

Private Sub cmdDeleteWord_Click()
    Dim i As Integer
    
    For i = CurrentWord + 1 To LearnDatabaseCount - 1
        LearnDatabase(i - 1).Question = LearnDatabase(i).Question
        LearnDatabase(i - 1).Answer = LearnDatabase(i).Answer
        LearnDatabase(i - 1).Image = LearnDatabase(i).Image
        LearnDatabase(i - 1).Sound = LearnDatabase(i).Sound
        LearnDatabase(i - 1).Language = LearnDatabase(i).Language
        DoEvents
    Next i
    
    LearnDatabase(LearnDatabaseCount - 1).Question = ""
    LearnDatabase(LearnDatabaseCount - 1).Answer = "???"
    LearnDatabase(LearnDatabaseCount - 1).Image = ""
    LearnDatabase(LearnDatabaseCount - 1).Sound = ""
    LearnDatabase(LearnDatabaseCount - 1).Language = ""

    LearnDatabaseCount = LearnDatabaseCount - 1
    
    lstMainList.RemoveItem (CurrentWord)
    
    If CurrentWord > LearnDatabaseCount - 1 Then CurrentWord = LearnDatabaseCount - 1
    
    FillMainList
    
    If LearnDatabaseCount > 0 Then
        lstMainList.Selected(CurrentWord) = True
        NewWordSelected
    Else
        cmdNewWord_Click
    End If
    
    FileChanged = True
    Form_Refresh
End Sub

Private Sub cmdSound_Browse_GotFocus()
    SetHighlight ehlSound, ehlFocus
End Sub

Private Sub cmdSound_Download_Click()
    If Len(txtRaw_Sound_Path.Text) > 0 And Len(txtRaw_Sound.Text) > 0 And ValidPath(txtRaw_Sound.Text) Then
        txtRaw_Sound_Path.Text = DownloadFile(txtRaw_Sound_Path.Text)
    ElseIf Len(txtRaw_Sound_Path.Text) <= 0 Then
        MsgBox "Kopiëer een internetadres naar het vakje 'Opzoeken'. (Bijvoorbeeld: http://www.audivididici.nl/test.wav)", vbInformation + vbOKOnly, "Audivididici"
    End If
End Sub

Private Sub cmdUndo_Click()
    NewWordSelected
    If FileChanged <> FileChangedBeforeWord Then
        FileChanged = FileChangedBeforeWord
        Form_Refresh
    End If
End Sub

Private Sub cmdUp_Click()
    Dim TempLearnType As LearnType
    Dim TempList As String
    
    If CurrentWord <= 0 Or LearnDatabaseCount <= 1 Then Exit Sub
    
    CheckSaveWord
    
    TempLearnType = LearnDatabase(CurrentWord)
    LearnDatabase(CurrentWord) = LearnDatabase(CurrentWord - 1)
    LearnDatabase(CurrentWord - 1) = TempLearnType
    
    TempList = lstMainList.List(CurrentWord)
    lstMainList.List(CurrentWord) = lstMainList.List(CurrentWord - 1)
    lstMainList.List(CurrentWord - 1) = TempList
    
    CurrentWord = CurrentWord - 1
    
    lstMainList.Selected(CurrentWord) = True
    NewWordSelected
    
    FileChanged = True
    Form_Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyN
                cmdNewWord_Click
            Case vbKeyS
                cmdSaveFile_Click
            Case vbKeyZ
                cmdUndo_Click
        End Select
    End If
End Sub

Private Sub Form_Load()
    Clip.StartClipboardViewer
    chkShowExample.value = Val(GetSetting(App.Title, "Edit", "chkShowExample.Value", IIf(Screen.Height > 10000, "1", "0")))
    chkShowExample_Click
    
    If Demo Then
        cmdSaveFile.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
    Form_Refresh
    RTFU_Question.HideMenu
    RTFU_Answer.HideMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim MsgboxResult As VbMsgBoxResult
    Dim SaveFileResult As Boolean
    
    CheckSaveWord
    
    If FileChanged Then
        If Demo Then
            MsgboxResult = MsgBox("Je wijzigingen worden niet opgeslagen! Toch afsluiten?", vbQuestion + vbYesNoCancel, "Audivididici Creator")
            
            If MsgboxResult = vbCancel Or MsgboxResult = vbNo Then
                Cancel = 1
                Exit Sub
            End If
        Else
            MsgboxResult = MsgBox("Wil je de wijzigingen opslaan?", vbQuestion + vbYesNoCancel, "Audivididici Creator")
            
            If MsgboxResult = vbCancel Then
                Cancel = 1
                Exit Sub
            Else
                If MsgboxResult = vbYes Then
                    SaveFileResult = SaveFile
                    Do While SaveFileResult = False
                        MsgboxResult = MsgBox("Weet je zeker dat je het bestand niet wil opslaan?", vbQuestion + vbYesNo, "Audivididici Creator")
                        If MsgboxResult = vbYes Then
                            SaveFileResult = True
                        Else
                            SaveFileResult = SaveFile
                        End If
                    Loop
                End If
            End If
        End If
    End If
    
    MMCSound.Command = "close"
    DelDir AppPath & "AVDTemp\"
    Dir AppPath
    
    frmMain.cmdEdit.Enabled = True
    frmMain.cmdStart.Enabled = True
    If Not Demo Then
        frmMain.cmdAdvanced.Enabled = True
        
        If GlobalFilePath <> "" And GlobalFileName <> "" Then
            frmMain.txtFile.Text = GlobalFilePath & GlobalFileName
            frmMain.NewFileSelected
        End If
    End If
    SaveSetting App.Title, "Edit", "chkShowExample.Value", chkShowExample.value
    
    Clip.EndClipboardViewer
End Sub

Private Sub lblAnswerBack_Click()
    On Error GoTo lblAnswerBack_Error
    cdl.Flags = cdlCCFullOpen + cdlCCRGBInit
    cdl.Color = shpAnswerBack.FillColor
    cdl.CancelError = True
    cdl.ShowColor
    
    txtRaw_Answer_Changing = True
    txtRaw_Answer.SelRTF = SetBackgroundColor(txtRaw_Answer.SelRTF, cdl.Color)
    shpAnswerBack.FillColor = cdl.Color
    txtRaw_Answer_Changing = False
    
lblAnswerBack_Error:
End Sub

Private Sub lblAnswerFront_Click()
    On Error GoTo lblAnswerFront_Error

    cdl.Flags = cdlCCFullOpen + cdlCCRGBInit
    cdl.Color = shpAnswerFront.FillColor
    cdl.CancelError = True
    cdl.ShowColor
    
    txtRaw_Answer_Changing = True
    txtRaw_Answer.SelColor = cdl.Color
    shpAnswerFront.FillColor = cdl.Color
    txtRaw_Answer_Changing = False
    
lblAnswerFront_Error:
End Sub

Private Sub lblAnswerHighlight_Click()
    txtRaw_Answer.SetFocus
End Sub

Private Sub lblAnswerHighlight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight ehlAnswer, ehlMouse
End Sub

Private Sub lblImageHighlight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight ehlImage, ehlMouse
End Sub

Private Sub lblQuestionBack_Click()
    On Error GoTo lblQuestionBack_Error
    cdl.Flags = cdlCCFullOpen + cdlCCRGBInit
    cdl.Color = shpQuestionBack.FillColor
    cdl.CancelError = True
    cdl.ShowColor
    
    txtRaw_Question_Changing = True
    txtRaw_Question.SelRTF = SetBackgroundColor(txtRaw_Question.SelRTF, cdl.Color)
    shpQuestionBack.FillColor = cdl.Color
    txtRaw_Question_Changing = False
    
lblQuestionBack_Error:
End Sub

Private Function SetBackgroundColor(TextRTF As String, BackgroundColor As OLE_COLOR) As String
    Dim ForeColor(999) As OLE_COLOR
    Dim i As Integer
    With tempRTF
        .TextRTF = TextRTF
        
        i = 0
        Do Until .SelStart >= Len(.Text) - 1 Or Len(.Text) = 0
            .SelStart = i
            .SelLength = 1
            ForeColor(i) = .SelColor
            i = i + 1
        Loop
        
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = BackgroundColor
        For i = 0 To 30
            .TextRTF = Replace(.TextRTF, "\highlight" & Trim$(Str(i)) & " ", "")
            If InStr(1, .TextRTF, "\highlight") < 1 Then Exit For
        Next i
        For i = 0 To 30
            .TextRTF = Replace(.TextRTF, "\highlight" & Trim$(Str(i)), "")
            If InStr(1, .TextRTF, "\highlight") < 1 Then Exit For
        Next i
        .TextRTF = Replace(.TextRTF, "\cf", "\highlight")
        
        i = 0
        Do Until .SelStart >= Len(.Text) - 1 Or Len(.Text) = 0
            .SelStart = i
            .SelLength = 1
            .SelColor = ForeColor(i)
            i = i + 1
        Loop
        
        .SelStart = 0
        .SelLength = Len(.Text)
        SetBackgroundColor = .SelRTF
    End With
End Function


Private Function GetBackgroundColor(TextRTF As String) As OLE_COLOR
    Dim i As Integer
    With tempRTF
        .TextRTF = TextRTF
        For i = 0 To 9
            .TextRTF = Replace(.TextRTF, "\cf" & Trim$(Str(i)), "")
        Next i
        If InStr(1, .TextRTF, "\highlight") <= 0 Then
            GetBackgroundColor = &HFFFFFF
            Exit Function
        End If
        .TextRTF = Replace(.TextRTF, "\highlight", "\cf")
        .SelStart = 0
        .SelLength = Len(.Text)
        
        GetBackgroundColor = .SelColor
    End With
End Function


Private Sub lblQuestionFront_Click()
    On Error GoTo lblQuestionFront_Error

    cdl.Flags = cdlCCFullOpen + cdlCCRGBInit
    cdl.Color = shpQuestionFront.FillColor
    cdl.CancelError = True
    cdl.ShowColor
    
    txtRaw_Question_Changing = True
    txtRaw_Question.SelColor = cdl.Color
    shpQuestionFront.FillColor = cdl.Color
    txtRaw_Question_Changing = False
    
lblQuestionFront_Error:
End Sub

Private Sub lblQuestionHighlight_Click()
    txtRaw_Question.SetFocus
End Sub

Private Sub lblQuestionHighlight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight ehlQuestion, ehlMouse
End Sub

Private Sub lblSoundHighlight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight ehlSound, ehlMouse
End Sub

Private Sub lstMainList_Click()
    Dim myFileChanged As Boolean
    Dim i As Integer
    
    CheckSaveWord
    
    For i = 0 To lstMainList.ListCount - 1
        If lstMainList.Selected(i) Then
            If CurrentWord = i Then Exit Sub
            
            CurrentWord = i
        End If
    Next i
    
    NewWordSelected
End Sub

Private Sub lstMainList_GotFocus()
    SetHighlight ehlMainList, ehlFocus
End Sub

Private Sub lstMainList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmdDeleteWord_Click
    End If
End Sub

Private Sub lstMainList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHighlight ehlMainList, ehlMouse
End Sub

Private Sub MMCSound_GotFocus()
    SetHighlight ehlSound, ehlFocus
End Sub

Private Sub MMCSound_PlayClick(Cancel As Integer)
    MMCSound.Command = "stop"
    MMCSound.Command = "close"
    MMCSound.Command = "open"
    If Not FExists(MMCSound.FileName) Then Exit Sub
    
    MMCSound.Command = "play"
End Sub

Private Sub MMCSound_RecordClick(Cancel As Integer)
    Recording = LearnDatabase(CurrentWord).Answer
    
    MMCSound.Command = "stop"
    MMCSound.Command = "close"
    
    DeleteFile AppPath & "TempRecord.wav"
    
    MMCSound.FileName = AppPath & "TempRecord.wav"
    MMCSound.Notify = False
    MMCSound.Wait = True
    MMCSound.Shareable = False
    MMCSound.DeviceType = "Waveaudio"
    MMCSound.RecordMode = mciRecordOverwrite
    MMCSound.Command = "open"
    MMCSound.Command = "record"
    
    OK_Sound = False
    If Not FileChanged Then
        FileChanged = True
        Form_Refresh
    End If
End Sub

Private Sub MMCSound_StopClick(Cancel As Integer)
    MMCSound.Command = "stop"
    MMCSound.Command = "save"
End Sub

Private Sub RTFU_Answer_ChangeRTF(SelRTF As String)
    txtRaw_Answer.SelRTF = SelRTF
    txtRaw_Answer.SelAlignment = rtfCenter
End Sub

Private Sub RTFU_Answer_CheckSigma()
    Dim OldSelStart As Integer
    Dim OldSelLength As Integer
    
    If Len(txtRaw_Answer.Text) > 0 And txtRaw_Answer.SelStart > 0 Then
        OldSelStart = txtRaw_Answer.SelStart
        OldSelLength = txtRaw_Answer.SelLength
        txtRaw_Answer.SelStart = txtRaw_Answer.SelStart - 1
        txtRaw_Answer.SelLength = 1
        RTFU_Answer.RTFValidate txtRaw_Answer.SelRTF, False
        txtRaw_Answer.SelStart = OldSelStart
        txtRaw_Answer.SelLength = OldSelLength
    End If
End Sub

Private Sub RTFU_Answer_Click()
    DoEvents
    RTFU_Answer.ShowMenu 0, 0
    DoEvents
End Sub

Private Sub RTFU_Answer_GotFocus()
    SetHighlight ehlAnswer, ehlFocus
End Sub

Private Sub RTFU_Answer_RTFSetFocus()
    On Error Resume Next
    txtRaw_Answer.SetFocus
End Sub

Private Sub RTFU_Question_ChangeRTF(SelRTF As String)
    DoEvents
    txtRaw_Question.SelRTF = SelRTF
    DoEvents
End Sub

Private Sub RTFU_Question_CheckSigma()
    Dim OldSelStart As Integer
    Dim OldSelLength As Integer
    
    If Len(txtRaw_Question.Text) > 0 And txtRaw_Question.SelStart > 0 Then
        OldSelStart = txtRaw_Question.SelStart
        OldSelLength = txtRaw_Question.SelLength
        txtRaw_Question.SelStart = txtRaw_Question.SelStart - 1
        txtRaw_Question.SelLength = 1
        RTFU_Question.RTFValidate txtRaw_Question.SelRTF, False
        txtRaw_Question.SelStart = OldSelStart
        txtRaw_Question.SelLength = OldSelLength
    End If
End Sub

Private Sub RTFU_Question_Click()
    RTFU_Question.ShowMenu 0, 0
End Sub

Private Sub RTFU_Question_GotFocus()
    SetHighlight ehlQuestion, ehlFocus
End Sub

Private Sub RTFU_Question_RTFSetFocus()
    On Error Resume Next
    txtRaw_Question.SetFocus
End Sub

Private Sub tmrSoundTracker_Timer()
    Dim PositieProcent As Integer
    If MMCSound.Position > 0 And MMCSound.Length > 0 Then
        PositieProcent = Int(100 / (MMCSound.Length / MMCSound.Position))
        StatusBar.Panels(3).Text = Streepjes(PositieProcent)
    Else
        StatusBar.Panels(3).Text = ""
    End If
End Sub

Private Function Streepjes(Procent As Integer) As String
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

Private Sub ClearDatabase()
    Dim i As Integer
    LearnDatabaseCount = 0
    For i = 0 To 999
        LearnDatabase(i).Question = ""
        LearnDatabase(i).Answer = ""
        LearnDatabase(i).Image = ""
        LearnDatabase(i).Sound = ""
        LearnDatabase(i).Language = ""
    Next i
End Sub

Public Function LoadFile(FilePath As String, FileName As String) As Boolean
    Dim MyFreeFile As Integer
    Dim MyLine As String
    Dim MyWordNumber As Integer
    Dim TempMax As Integer
    Dim RetCode As Long
    
    MyFreeFile = FreeFile()
    MyWordNumber = 0
    
    LoadFile = True
    
    If Right$(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    
    If Not FExists(FilePath & FileName) Then
        MsgBox "Bestand niet gevonden!", vbCritical, "Audivididici Creator"
        LoadFile = False
        Unload Me
        Exit Function
    End If
    
    GlobalFilePath = FilePath
    GlobalFileName = FileName
    ClearDatabase
    
    Load frmSplash
    frmSplash.SetMaximumAndValue 1, 0
    frmSplash.Show
    
    MMCSound.Command = "close"
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
            LearnDatabase(MyWordNumber).Question = MyLine
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            
            Line Input #MyFreeFile, MyLine
            LearnDatabase(MyWordNumber).Answer = MyLine
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            
            Line Input #MyFreeFile, MyLine
            MyLine = Trim$(MyLine)
            LearnDatabase(MyWordNumber).Image = MyLine
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            
            Line Input #MyFreeFile, MyLine
            MyLine = Trim$(MyLine)
            LearnDatabase(MyWordNumber).Sound = MyLine
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            
            Line Input #MyFreeFile, MyLine
            MyLine = Trim$(MyLine)
            LearnDatabase(MyWordNumber).Language = MyLine

DoEinde:
            
            MyWordNumber = MyWordNumber + 1
            
            frmSplash.SetValue MyWordNumber + 1
            DoEvents
        Loop
    Close #MyFreeFile
    LearnDatabaseCount = MyWordNumber
    
    OK_Question = True
    OK_Answer = True
    OK_Image = True
    OK_Sound = True
    
    FillMainList
    lstMainList.Selected(0) = True
    NewWordSelected
    
    FileChanged = False
    Form_Refresh
    
    Unload frmSplash
End Function

Private Sub FillMainList()
    Dim i As Integer
    Dim SelectedOne As Integer
    Dim myQuestion As String
    Dim myAnswer As String
    
    i = 0
    SelectedOne = -1
    Do While i < lstMainList.ListCount And SelectedOne = -1
        If lstMainList.Selected(i) = True Then
            SelectedOne = i
        End If
        i = i + 1
    Loop
    
    lstMainList.Clear
    For i = 0 To LearnDatabaseCount - 1
        myQuestion = Replace(NormalText(LearnDatabase(i).Question), "?", "")
        myAnswer = Replace(NormalText(LearnDatabase(i).Answer), "?", "")
                
        lstMainList.AddItem ListMarkup(i + 1, myQuestion, myAnswer, 32)
        
        If i = SelectedOne Then
            lstMainList.Selected(i) = True
        Else
            lstMainList.Selected(i) = False
        End If
    Next
End Sub

Private Function TrimNicely(Text As String, Length As Integer)
    If Len(Text) <= Length Then
        TrimNicely = Text
    Else
        TrimNicely = Trim$(Left$(Text, Length - 3)) & "..."
    End If
End Function

Private Sub txtRaw_Answer_Change()
    OK_Answer = False
    If Not FileChanged Then
        FileChanged = True
        Form_Refresh
    End If
End Sub

Private Sub txtRaw_Answer_GotFocus()
    If txtRaw_Answer.Text = "???" Then
        txtRaw_Answer.SelStart = 0
        txtRaw_Answer.SelLength = Len(txtRaw_Answer)
    End If
    SetHighlight ehlAnswer, ehlFocus
End Sub

Private Sub txtRaw_Answer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdNewWord.SetFocus
    End If
End Sub

Private Sub txtRaw_Answer_LostFocus()
    Dim TempSelStart As Integer
    Dim TempSelLength As Integer
    
    If Not txtRaw_Answer_Changing And Not OK_Answer Then
        TempSelStart = txtRaw_Answer.SelStart
        TempSelLength = txtRaw_Answer.SelLength
        txtRaw_Answer.SelStart = 0
        txtRaw_Answer.SelLength = Len(txtRaw_Answer.Text)
        txtRaw_Answer.SelFontName = "Gentium"
        txtRaw_Answer.SelFontSize = 18
        txtRaw_Answer.SelAlignment = rtfCenter
        txtRaw_Answer.SelStart = TempSelStart
        txtRaw_Answer.SelLength = TempSelLength
    End If
    
    MakePath False
End Sub

Private Sub txtRaw_Answer_SelChange()
    On Error Resume Next
    If Not txtRaw_Answer_Changing Then
        txtRaw_Answer_SelChanging = True
        chkAnswerBold.value = IIf(txtRaw_Answer.SelBold, 1, 0)
        chkAnswerItalic.value = IIf(txtRaw_Answer.SelItalic, 1, 0)
        chkAnswerUnderline.value = IIf(txtRaw_Answer.SelUnderline, 1, 0)
        chkAnswerStrikeThru.value = IIf(txtRaw_Answer.SelStrikeThru, 1, 0)
        shpAnswerFront.FillColor = txtRaw_Answer.SelColor
        shpAnswerBack.FillColor = GetBackgroundColor(txtRaw_Answer.SelRTF)
        txtRaw_Answer_SelChanging = False
    End If
End Sub

Private Sub txtRaw_Image_Change()
    SetRawImage
End Sub

Private Sub txtRaw_Image_GotFocus()
    SetHighlight ehlImage, ehlFocus
End Sub

Private Sub txtRaw_Image_Path_Change()
    SetRawImage
End Sub

Private Sub txtRaw_Image_Path_GotFocus()
    SetHighlight ehlImage, ehlFocus
End Sub

Private Sub txtRaw_Question_Change()
    OK_Question = False
    If Not FileChanged Then
        FileChanged = True
        Form_Refresh
    End If
End Sub

Private Function CleanRTF(TextRTF As String) As String
    Dim retVal, NewString As String
    
    NewString = TextRTF
    
    Do Until NewString = retVal
        retVal = NewString
        NewString = Replace(retVal, "\par \par", "\par")
        NewString = Replace(retVal, vbCrLf, "")
    Loop
    retVal = NewString
    
    CleanRTF = retVal
End Function

Private Sub MakePath(Question As Boolean)
    Dim Length, i As Integer
    Dim Ext, myAnswer As String
    Dim Number As String
    
    If Question Then
        myAnswer = Trim$(NormalText(CleanRTF(txtRaw_Question.TextRTF)))
    Else
        myAnswer = Trim$(NormalText(CleanRTF(txtRaw_Answer.TextRTF)))
    End If
    myAnswer = Replace(myAnswer, " ", "_")
    myAnswer = Replace(myAnswer, "?", "")
    
    i = Len(myAnswer & "\")
    Length = 0
    Do Until Length = i - 1 Or i = 0
        Length = i - 1
        i = InStr(1, Left$(myAnswer & "\", Length), " ")
        If i = 0 Or i > 10 Then
            i = InStr(1, Left$(myAnswer & "\", Length), ",")
            If i = 0 Or i > 10 Then
                i = InStr(1, Left$(myAnswer & "\", Length), ";")
                If i = 0 Or i > 10 Then
                    i = InStr(1, Left$(myAnswer & "\", Length), ":")
                    If i = 0 Or i > 10 Then
                        i = InStr(1, Left$(myAnswer & "\", Length), "!")
                        If i = 0 Or i > 10 Then
                            i = InStr(1, Left$(myAnswer & "\", Length), ".")
                            If i = 0 Or i > 10 Then
                                i = InStr(1, Left$(myAnswer & "\", Length), "\")
                                If i = 0 Or i > 10 Then
                                    i = InStr(1, Left$(myAnswer & "\", Length), "/")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop
    
    If Length <= 2 Then
        If Question Then
            myAnswer = Trim$(Str$(CurrentWord + 1))
            Length = Len(myAnswer)
        Else
            MakePath True
            Exit Sub
        End If
    End If
    
    Ext = ".jpg"
    If Len(txtRaw_Image_Path.Text) > 6 Then
        If Left$(Right$(txtRaw_Image_Path.Text, 4), 1) = "." Then
            Ext = Right$(txtRaw_Image_Path.Text, 4)
        ElseIf Left$(Right$(txtRaw_Image_Path.Text, 5), 1) = "." Then
            Ext = Right$(txtRaw_Image_Path.Text, 5)
        End If
    End If

    If LearnDatabase(CurrentWord).Image = "" Then
        If Length > 1 Then
            Number = ""
            
            Do While FExists(AppPath & "AVDTemp\" & Left$(myAnswer, Length) & Number & Ext)
                Number = "_" & Trim(Str$(Int(Val(Number)) + 1))
            Loop
            
            txtRaw_Image.Text = Left$(myAnswer, Length) & Number & Ext
        Else
            txtRaw_Image.Text = ""
        End If
    End If
    
    Ext = ".wav"
    If Len(txtRaw_Sound_Path.Text) > 6 Then
        If Left$(Right$(txtRaw_Sound_Path.Text, 4), 1) = "." Then
            Ext = Right$(txtRaw_Sound_Path.Text, 4)
        End If
    End If
    
    If LearnDatabase(CurrentWord).Sound = "" Then
        If Length > 1 Then
            Number = ""
            
            Do While FExists(AppPath & "AVDTemp\" & Left$(myAnswer, Length) & Number & Ext)
                Number = "_" & Trim(Str$(Int(Val(Number)) + 1))
            Loop
            
            txtRaw_Sound.Text = Left$(myAnswer, Length) & Number & Ext
        Else
            txtRaw_Sound.Text = ""
        End If
    End If
End Sub

Private Sub SetRawImage()
    On Error GoTo SetRawImageError
    
    With imgRaw_Image
        OK_Image = False
        
        If txtRaw_Image.Text = LearnDatabase(CurrentWord).Image And FExists(AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Image) Then
            OK_Image = True
        End If
        
        If FExists(txtRaw_Image_Path.Text) Then
            .Picture = LoadPicture(txtRaw_Image_Path.Text)
            .Visible = True
        Else
            If FExists(AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Image) Then
                .Picture = LoadPicture(AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Image)
                .Visible = True
            Else
                If FExists(AppPath & "AVDTemp\" & txtRaw_Image.Text) Then
                    .Picture = LoadPicture(AppPath & "AVDTemp\" & txtRaw_Image.Text)
                    .Visible = True
                Else
                    .Visible = False
                End If
            End If
        End If
        
        If .Visible Then
            If .Picture.Width > 0 And .Picture.Width > 0 Then
                If .Picture.Width > .Picture.Height Then
                    .Width = imgRaw_Image_Width
                    .Height = (.Picture.Height / .Picture.Width) * imgRaw_Image_Height
                Else
                    .Width = (.Picture.Width / .Picture.Height) * imgRaw_Image_Width
                    .Height = imgRaw_Image_Height
                End If
                
                .Top = imgRaw_Image_Top + imgRaw_Image_Height / 2 - .Height / 2
                .Left = imgRaw_Image_Left + imgRaw_Image_Width / 2 - .Width / 2
            Else
                .Visible = False
            End If
        End If
    End With
    
    With imgWord_Image
        If imgRaw_Image.Visible Then
            .Picture = imgRaw_Image.Picture
            .Visible = True
            If .Picture.Width > .Picture.Height Then
                .Width = imgWord_Image_Width
                .Height = (.Picture.Height / .Picture.Width) * imgWord_Image_Height
            Else
                .Width = (.Picture.Width / .Picture.Height) * imgWord_Image_Width
                .Height = imgWord_Image_Height
            End If
            
            .Top = imgWord_Image_Top + imgWord_Image_Height / 2 - .Height / 2
            .Left = imgWord_Image_Left + imgWord_Image_Width / 2 - .Width / 2
        Else
            .Visible = False
        End If
    End With
    
    If Not FileChanged And Not OK_Image Then
        FileChanged = True
        Form_Refresh
    End If
    
    Exit Sub
SetRawImageError:
    MsgBox "Je hebt geen geldig bestand geselecteerd!", vbCritical + vbOKOnly, "Audivididici Creator"
    txtRaw_Image.Text = ""
End Sub

Private Sub SetRawSound()
    On Error GoTo SetRawSoundError
    
    If Recording = LearnDatabase(CurrentWord).Answer And LearnDatabase(CurrentWord).Answer <> "" Then Exit Sub
    
    OK_Sound = False
    
    If txtRaw_Sound.Text = LearnDatabase(CurrentWord).Sound And FExists(AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Sound) Then
        OK_Sound = True
    End If
    
    If FExists(txtRaw_Sound_Path.Text) Then
        If MMCSound.FileName = txtRaw_Sound_Path.Text Then Exit Sub
        MMCSound.FileName = txtRaw_Sound_Path.Text
    Else
        If FExists(AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Sound) Then
            If MMCSound.FileName = AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Sound Then Exit Sub
            MMCSound.FileName = AppPath & "AVDTemp\" & LearnDatabase(CurrentWord).Sound
        Else
            If FExists(AppPath & "AVDTemp\" & txtRaw_Sound.Text) Then
                OK_Sound = True
                If MMCSound.FileName = AppPath & "AVDTemp\" & txtRaw_Sound.Text Then Exit Sub
                MMCSound.FileName = AppPath & "AVDTemp\" & txtRaw_Sound.Text
            Else
                MMCSound.FileName = AppPath & "TempRecord.wav"
            End If
        End If
    End If
    MMCSound.Command = "close"
    MMCSound.Notify = False
    MMCSound.Wait = True
    MMCSound.Shareable = False
    MMCSound.DeviceType = "Waveaudio"
    MMCSound.Command = "open"
    
    If Not FileChanged And Not OK_Sound Then
        FileChanged = True
        Form_Refresh
    End If

    Exit Sub
SetRawSoundError:
    MsgBox "Je hebt geen geldig bestand geselecteerd!", vbCritical + vbOKOnly, "Audivididici Creator"
    txtRaw_Sound.Text = ""
End Sub

Private Function ValidPath(Path As String) As Boolean
    ValidPath = (InStr(1, Path, "/.") = 0)
    If (InStr(1, Path, "?") > 0) Then
        ValidPath = False
    End If
End Function

Private Sub txtRaw_Question_GotFocus()
    SetHighlight ehlQuestion, ehlFocus
End Sub

Private Sub txtRaw_Question_KeyDown(KeyCode As Integer, Shift As Integer)
    RTFU_Question.RTFKeyDown KeyCode, Shift
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub txtRaw_Question_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtRaw_Answer.SetFocus
    End If
End Sub

Private Sub txtRaw_Question_KeyUp(KeyCode As Integer, Shift As Integer)
    RTFU_Question.RTFKeyUp KeyCode, Shift
End Sub

Private Sub txtRaw_Answer_KeyDown(KeyCode As Integer, Shift As Integer)
    RTFU_Answer.RTFKeyDown KeyCode, Shift
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub txtRaw_Answer_KeyUp(KeyCode As Integer, Shift As Integer)
    RTFU_Answer.RTFKeyUp KeyCode, Shift
End Sub

Private Sub txtRaw_Question_LostFocus()
    Dim TempSelStart As Integer
    Dim TempSelLength As Integer
    
    If Not txtRaw_Question_Changing And Not OK_Question Then
        TempSelStart = txtRaw_Question.SelStart
        TempSelLength = txtRaw_Question.SelLength
        txtRaw_Question.SelStart = 0
        txtRaw_Question.SelLength = Len(txtRaw_Question.Text)
        txtRaw_Question.SelFontName = "Gentium"
        txtRaw_Question.SelFontSize = 18
        txtRaw_Question.SelAlignment = rtfCenter
        txtRaw_Question.SelStart = TempSelStart
        txtRaw_Question.SelLength = TempSelLength
    End If
    
    MakePath False
End Sub

Private Sub txtRaw_Question_SelChange()
    On Error Resume Next
    If Not txtRaw_Question_Changing Then
        txtRaw_Question_SelChanging = True
        chkQuestionBold.value = IIf(txtRaw_Question.SelBold, 1, 0)
        chkQuestionItalic.value = IIf(txtRaw_Question.SelItalic, 1, 0)
        chkQuestionUnderline.value = IIf(txtRaw_Question.SelUnderline, 1, 0)
        chkQuestionStrikeThru.value = IIf(txtRaw_Question.SelStrikeThru, 1, 0)
        shpQuestionFront.FillColor = txtRaw_Question.SelColor
        shpQuestionBack.FillColor = GetBackgroundColor(txtRaw_Question.SelRTF)
        txtRaw_Question_SelChanging = False
    End If
End Sub

Private Function NormalText(MyString As String) As String
    tempRTF.TextRTF = MyString
    If Len(tempRTF.Text) < 2 And Len(MyString) >= 2 And Left(MyString, 5) <> "{\rtf" Then
        NormalText = MyString
    Else
        NormalText = tempRTF.Text
    End If
End Function

Private Sub SetHighlight(myControl As EditHighlightConstants, myType As EditHightlightTypes)
    Dim OldControl As EditHighlightConstants
    Dim OtherControl As EditHighlightConstants
    
    OldControl = IIf(myType = ehlFocus, FocusHighlight, MouseHighlight)
    OtherControl = IIf(myType = ehlMouse, FocusHighlight, MouseHighlight)
    
    If myControl = OldControl Then Exit Sub
    
    If OtherControl <> OldControl Then
        Select Case OldControl
            Case ehlQuestion
                shpQuestion.FillColor = &HFFFFFF
                chkQuestionBold.Visible = False
                chkQuestionItalic.Visible = False
                chkQuestionUnderline.Visible = False
                chkQuestionStrikeThru.Visible = False
                shpQuestionFront.Visible = False
                shpQuestionBack.Visible = False
                lblQuestionFront.Visible = False
                lblQuestionBack.Visible = False
                cmbQuestion.Visible = False
                RTFU_Question.Visible = False
                'txtRaw_Question.ScrollBars = rtfNone
                If myType = ehlFocus Then
                    txtWord_Question.TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}}\viewkind4\pard\f0\fs38\qc " & CleanRTF(txtRaw_Question.TextRTF) & "\par}"
                End If
            Case ehlAnswer
                shpAnswer.FillColor = &HFFFFFF
                chkAnswerBold.Visible = False
                chkAnswerItalic.Visible = False
                chkAnswerUnderline.Visible = False
                chkAnswerStrikeThru.Visible = False
                shpAnswerFront.Visible = False
                shpAnswerBack.Visible = False
                lblAnswerFront.Visible = False
                lblAnswerBack.Visible = False
                cmbAnswer.Visible = False
                RTFU_Answer.Visible = False
                'txtRaw_Answer.ScrollBars = rtfNone
                If myType = ehlFocus Then
                    txtWord_Answer.TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}}\viewkind4\uc1\pard\f0\fs32\qc " & CleanRTF(txtRaw_Answer.TextRTF) & "\par}"
                End If
            Case ehlImage
                shpImage.FillColor = &HFFFFFF
            Case ehlSound
                shpSound.FillColor = &HFFFFFF
            Case ehlUndo
                cmdUndo.BackColor = &HFFFFFF
            Case ehlNewWord
                cmdNewWord.BackColor = &HFFFFFF
            Case ehlDeleteWord
                cmdDeleteWord.BackColor = &HFFFFFF
            Case ehlMainList
                lstMainList.BackColor = &HFFFFFF
            Case ehlsavefile
                cmdSaveFile.BackColor = &HFFFFFF
        End Select
    End If
    
    If OtherControl <> myControl Then
        Select Case myControl
            Case ehlQuestion
                shpQuestion.FillColor = &HC0FFFF
                chkQuestionBold.Visible = True
                chkQuestionItalic.Visible = True
                chkQuestionUnderline.Visible = True
                chkQuestionStrikeThru.Visible = True
                shpQuestionFront.Visible = True
                shpQuestionBack.Visible = True
                lblQuestionFront.Visible = True
                lblQuestionBack.Visible = True
                cmbQuestion.Visible = True
                RTFU_Question.Visible = True
                'txtRaw_Question.ScrollBars = rtfHorizontal
            Case ehlAnswer
                shpAnswer.FillColor = &HC0FFFF
                chkAnswerBold.Visible = True
                chkAnswerItalic.Visible = True
                chkAnswerUnderline.Visible = True
                chkAnswerStrikeThru.Visible = True
                shpAnswerFront.Visible = True
                shpAnswerBack.Visible = True
                lblAnswerFront.Visible = True
                lblAnswerBack.Visible = True
                cmbAnswer.Visible = True
                RTFU_Answer.Visible = True
                'txtRaw_Answer.ScrollBars = rtfHorizontal
            Case ehlImage
                shpImage.FillColor = &HC0FFFF
            Case ehlSound
                shpSound.FillColor = &HC0FFFF
            Case ehlUndo
                cmdUndo.BackColor = &HC0FFFF
            Case ehlNewWord
                cmdNewWord.BackColor = &HC0FFFF
            Case ehlDeleteWord
                cmdDeleteWord.BackColor = &HC0FFFF
            Case ehlMainList
                lstMainList.BackColor = &HC0FFFF
            Case ehlsavefile
                cmdSaveFile.BackColor = &HC0FFFF
        End Select
    End If
    
    If myType = ehlFocus Then
        FocusHighlight = myControl
    ElseIf myType = ehlMouse Then
        MouseHighlight = myControl
    End If
End Sub

Private Sub txtRaw_Sound_Change()
    SetRawSound
End Sub

Private Sub txtRaw_Sound_GotFocus()
    SetHighlight ehlSound, ehlFocus
End Sub

Private Sub txtRaw_Sound_Path_Change()
    SetRawSound
End Sub

Private Sub txtRaw_Sound_Path_GotFocus()
    SetHighlight ehlSound, ehlFocus
End Sub

Public Sub SetCompression(myCompression As Integer)
    If myCompression < 0 Or myCompression > 9 Then
        myCompression = 5
    End If
    
    Compression = myCompression
End Sub
 
Private Function DownloadFile(URL As String) As String
    On Error GoTo DownloadErr
    Dim szFileName As String
    
    fraDownload.Visible = True
    
    DoEvents
    
    szFileName = Space$(300)
    
    DownloadFile = ""
    
    If URLDownloadToCacheFile(0, URL, szFileName, Len(szFileName), 0, 0) = 0 Then
        DownloadFile = Trim(szFileName)
    End If
    
    If InStr(1, DownloadFile, ".htm") > 0 Then
        If InStr(1, LCase(URL), "wiki", vbTextCompare) > 0 Then
            MsgBox "Hint: bij Wikipedia moet je een aantal keer op de afbeelding klikken voor je deze kunt kopieeren!", vbInformation + vbOKOnly, "Audivididici Creator"
        Else
            MsgBox "Hint: wellicht moet je eerst op de afbeelding klikken alvorens je iets kan kopieeren.", vbInformation + vbOKOnly, "Audivididici Creator"
        End If
        
        DownloadFile = ""
    End If
    
    fraDownload.Visible = False
     
    Exit Function
DownloadErr:

    fraDownload.Visible = False
End Function

Sub CheckSaveWord()
    If (OK_Question = False) Or (OK_Answer = False) Or (OK_Image = False And imgRaw_Image.Visible = True) Or (OK_Sound = False And FExists(MMCSound.FileName) = True) Then
        SaveWord
    End If
End Sub

