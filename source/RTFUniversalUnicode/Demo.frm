VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Universal Unicode ActiveX Control Demo"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7440
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin RTFUUDemo.RTFUniversalUnicode RTFUniversalUnicode1 
      Height          =   1935
      Left            =   5400
      TabIndex        =   6
      Top             =   315
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3413
   End
   Begin VB.OptionButton optGreek 
      Caption         =   "Griekse indeling"
      Height          =   195
      Left            =   5400
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.OptionButton optLatin 
      Caption         =   "Latijnse indeling"
      Height          =   195
      Left            =   5400
      TabIndex        =   4
      Top             =   1800
      Value           =   -1  'True
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Demo.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Gentium"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdRTF 
      Caption         =   "Troubleshoot RTF"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   $"Demo.frx":007C
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   $"Demo.frx":0111
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRTF_Click()
    Form1.Show
End Sub

Private Sub Form_Load()
    RTFUniversalUnicode1.HideMenu
End Sub

Private Sub optGreek_Click()
    RTFUniversalUnicode1.SetKeyboard kbdGreek
    RichTextBox1.SetFocus
End Sub

Private Sub optLatin_Click()
    RTFUniversalUnicode1.SetKeyboard kbdLatin
    RichTextBox1.SetFocus
End Sub

Private Sub RichTextBox1_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    RTFUniversalUnicode1.RTFKeyDown KeyCode, Shift
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    RTFUniversalUnicode1.RTFKeyUp KeyCode, Shift
End Sub

Private Sub RTFUniversalUnicode1_ChangeRTF(SelRTF As String)
    RichTextBox1.SelRTF = SelRTF
End Sub

Private Sub RTFUniversalUnicode1_CheckSigma()
    Dim OldSelStart As Integer
    Dim OldSelLength As Integer
    
    If Len(RichTextBox1.Text) > 0 And RichTextBox1.SelStart > 0 Then
        OldSelStart = RichTextBox1.SelStart
        OldSelLength = RichTextBox1.SelLength
        RichTextBox1.SelStart = RichTextBox1.SelStart - 1
        RichTextBox1.SelLength = 1
        RTFUniversalUnicode1.RTFValidate RichTextBox1.SelRTF, False
        RichTextBox1.SelStart = OldSelStart
        RichTextBox1.SelLength = OldSelLength
    End If
End Sub

Private Sub RTFUniversalUnicode1_Click()
    RTFUniversalUnicode1.ShowMenu 0, 0
End Sub

Private Sub RTFUniversalUnicode1_RTFSetFocus()
    On Error Resume Next
    RichTextBox1.SetFocus
End Sub
