VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmResults 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audivididici"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6180
   Icon            =   "frmResults.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6180
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optTop3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Top 3"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3263
      TabIndex        =   16
      Top             =   3240
      Width           =   855
   End
   Begin VB.OptionButton optProgress 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Voortgang"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2063
      TabIndex        =   15
      Top             =   3240
      Value           =   -1  'True
      Width           =   1080
   End
   Begin VB.Frame fraStats 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3135
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.OptionButton optTop3Choise 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Beste vooruitgang"
         ForeColor       =   &H000080FF&
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   14
         Top             =   200
         Width           =   1215
      End
      Begin VB.OptionButton optTop3Choise 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Moeilijkste woorden in totaal"
         ForeColor       =   &H000080FF&
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   13
         Top             =   200
         Width           =   1575
      End
      Begin VB.OptionButton optTop3Choise 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Moeilijkste woorden deze keer"
         ForeColor       =   &H000080FF&
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   12
         Top             =   200
         Width           =   1815
      End
      Begin RichTextLib.RichTextBox txtTop3 
         Height          =   735
         Index           =   0
         Left            =   600
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1296
         _Version        =   393217
         ReadOnly        =   -1  'True
         MousePointer    =   1
         Appearance      =   0
         TextRTF         =   $"frmResults.frx":08CA
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
      Begin RichTextLib.RichTextBox txtTop3 
         Height          =   735
         Index           =   1
         Left            =   600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1296
         _Version        =   393217
         ReadOnly        =   -1  'True
         MousePointer    =   1
         Appearance      =   0
         TextRTF         =   $"frmResults.frx":0943
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
      Begin RichTextLib.RichTextBox txtTop3 
         Height          =   735
         Index           =   2
         Left            =   600
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2400
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1296
         _Version        =   393217
         ReadOnly        =   -1  'True
         MousePointer    =   1
         Appearance      =   0
         TextRTF         =   $"frmResults.frx":09BC
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
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   2500
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1660
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   820
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Top 3"
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
         Left            =   120
         TabIndex        =   8
         Top             =   200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   6135
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Default         =   -1  'True
         Height          =   615
         Left            =   2423
         Picture         =   "frmResults.frx":0A35
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Timer tmrCount 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   480
         Top             =   600
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Je had 70% van de antwoorden in één keer goed."
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
         Height          =   735
         Left            =   143
         TabIndex        =   2
         Top             =   120
         Width           =   5895
      End
   End
   Begin MSChart20Lib.MSChart Chart 
      DragMode        =   1  'Automatic
      Height          =   3735
      Left            =   -360
      OleObjectBlob   =   "frmResults.frx":0B7D
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -120
      Width           =   6615
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Scores(5) As Integer
Private Times(5) As String
Private ScoresCount As Integer

Private Top3_Hardest(2) As String
Private Top3_Hardest_Total(2) As String
Private Top3_Best_Progress(2) As String


Public Sub ClearScores()
    Dim i As Integer
    
    For i = 0 To 5
        Scores(i) = 0
        Times(i) = ""
    Next i
    
    ScoresCount = 0
    
    ChartRefresh
End Sub

Public Sub NewScore(myScore As Integer, Optional myTime As String = "")
    Dim i As Integer
    
    If ScoresCount < 6 Then
        ScoresCount = ScoresCount + 1
    End If
    
    For i = 5 To 1 Step -1
        Scores(i) = Scores(i - 1)
        Times(i) = Times(i - 1)
    Next i
    
    Scores(0) = myScore
    
    If myTime = "" Then
        Times(0) = Hour(Now()) & ":" & IIf(Minute(Now()) < 10, "0", "") & Minute(Now())
    Else
        Times(0) = myTime
    End If
    
    lblScore.Visible = False
    optProgress.value = True
    optTop3Choise(0).value = True
    
    ChartRefresh
End Sub

Public Sub ChartRefresh()
    Dim i As Integer
    Dim j As Integer
    
    With Chart
        .RowCount = ScoresCount
        
        j = 1
        
        If ScoresCount > 1 Then
            For i = ScoresCount - 1 To 1 Step -1
                .Row = j
                .Data = Scores(i)
                .RowLabel = Times(i)
                
                j = j + 1
            Next i
        End If
        
        If ScoresCount > 0 Then
            .Row = j
            .RowLabel = Times(0)
            
            If lblScore.Visible Then
                .Data = Scores(0)
            Else
                .Data = 0
                tmrCount.Enabled = True
            End If
        End If
    End With
End Sub

Private Sub cmdOk_Click()
    Me.Hide
    frmWords.StartAgain
End Sub

Private Sub optProgress_Click()
    fraStats.Visible = False
End Sub

Private Sub optTop3_Click()
    fraStats.Visible = True
End Sub

Public Sub optTop3Choise_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 2
        Select Case Index
            Case 0
                txtTop3(i).TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}{\f1\fnil\fprq2\fcharset161 Gentium;}{\f2\fnil\fprq2\fcharset238 Gentium;}}\viewkind4\pard\f0\fs38\qc " & Top3_Hardest(i) & "\par}"
            Case 1
                txtTop3(i).TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}{\f1\fnil\fprq2\fcharset161 Gentium;}{\f2\fnil\fprq2\fcharset238 Gentium;}}\viewkind4\pard\f0\fs38\qc " & Top3_Hardest_Total(i) & "\par}"
            Case 2
                txtTop3(i).TextRTF = "{\rtf1{\fonttbl{\f0 Gentium;}{\f1\fnil\fprq2\fcharset161 Gentium;}{\f2\fnil\fprq2\fcharset238 Gentium;}}\viewkind4\pard\f0\fs38\qc " & Top3_Best_Progress(i) & "\par}"
        End Select
    Next i
End Sub

Private Sub tmrCount_Timer()
    If (Chart.Data) = Scores(0) Then
        tmrCount.Enabled = False
        lblScore.Caption = "Je had " & Scores(0) & "% van de antwoorden in één keer goed."
        lblScore.Visible = True
    End If
    
    Chart.Data = Chart.Data + 1
End Sub

Public Sub SetTop3(Top3_Hardest_1 As String, Top3_Hardest_2 As String, Top3_Hardest_3 As String, Top3_Hardest_Total_1 As String, Top3_Hardest_Total_2 As String, Top3_Hardest_Total_3 As String, Top3_Best_Progress_1 As String, Top3_Best_Progress_2 As String, Top3_Best_Progress_3 As String, Top3_Visible As Boolean)
    Top3_Hardest(0) = Top3_Hardest_1
    Top3_Hardest(1) = Top3_Hardest_2
    Top3_Hardest(2) = Top3_Hardest_3
    Top3_Hardest_Total(0) = Top3_Hardest_Total_1
    Top3_Hardest_Total(1) = Top3_Hardest_Total_2
    Top3_Hardest_Total(2) = Top3_Hardest_Total_3
    Top3_Best_Progress(0) = Top3_Best_Progress_1
    Top3_Best_Progress(1) = Top3_Best_Progress_2
    Top3_Best_Progress(2) = Top3_Best_Progress_3
    
    optProgress.Visible = Top3_Visible
    optTop3.Visible = Top3_Visible
End Sub
