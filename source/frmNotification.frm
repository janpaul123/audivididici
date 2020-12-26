VERSION 5.00
Begin VB.Form frmNotification 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audivididici Demo"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   Icon            =   "frmNotification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdWebsite 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Website"
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
      Left            =   2400
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   480
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ok"
      Default         =   -1  'True
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
      Left            =   120
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblNotification2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "U kunt hiermee alleen voorbeeld-bestanden openen en geen wijzigingen opslaan."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label lblNotification1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dit is de demo-versie van Audivididici"
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
      Height          =   855
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdOk.BackColor = &HC0FFFF
End Sub

Private Sub cmdWebsite_Click()
    ShellExecute 0&, vbNullString, "http://www.audivididici.nl", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub cmdWebsite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdWebsite.BackColor = &HC0FFFF
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdOk.BackColor = &HFFFFFF
    cmdWebsite.BackColor = &HFFFFFF
End Sub

Private Sub Timer1_Timer()
    cmdOk.Enabled = True
    cmdOk.Default = True
    Timer1.Enabled = False
    cmdOk.SetFocus
End Sub
