VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "RTFTest.frx":0000
      Top             =   2640
      Width           =   6255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   135
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "RTFTest.frx":000C
      Top             =   1200
      Width           =   6255
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1508
      _Version        =   393217
      TextRTF         =   $"RTFTest.frx":0012
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RichTextBox1_Change()
    If Option1(0).Value Then
        Text1.Text = RichTextBox1.TextRTF
        Text2.Text = RichTextBox1.SelRTF
    End If
End Sub

Private Sub RichTextBox1_SelChange()
    RichTextBox1_Change
End Sub

Private Sub Text1_Change()
    If Option1(1).Value Then RichTextBox1.TextRTF = Text1.Text
End Sub

Private Sub Text2_Change()
    If Option1(2).Value Then RichTextBox1.SelRTF = Text2.Text
End Sub
