VERSION 5.00
Begin VB.Form frmExample 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audivididici"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
   Icon            =   "frmExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ok"
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
      Left            =   1260
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton optFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aanhef Odyssee"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   480
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.Label lblCompression 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Voorbeeld van scanderen in het Grieks."
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Selecteer hier een voorbeeldbestand:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00037BE9&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    If optFile(0).value = True Then
        If FExists(AppPath & "Odyssee.dem") Then
            frmMain.DemoFile = AppPath & "Odyssee.dem"
            frmMain.txtFile = "Aanhef Odyssee"
        Else
            MsgBox "Voorbeeldbestand niet gevonden!", vbCritical + vbOKOnly, "Audivididici"
        End If
    End If
    
    Unload Me
End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdOk.BackColor = &HC0FFFF
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdOk.BackColor = &HFFFFFF
End Sub

