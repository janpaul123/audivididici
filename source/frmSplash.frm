VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Laden - Audivididici"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   Begin VB.Shape shpHide 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   360
      Top             =   1320
      Width           =   6255
   End
   Begin VB.Label lblDidiciText 
      BackStyle       =   0  'Transparent
      Caption         =   "ik leerde"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   4560
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblVidiText 
      BackStyle       =   0  'Transparent
      Caption         =   "ik zag"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblAudiviText 
      BackStyle       =   0  'Transparent
      Caption         =   "ik hoorde"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblVidi 
      BackStyle       =   0  'Transparent
      Caption         =   "vidi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1335
      Left            =   2685
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblDidici 
      BackStyle       =   0  'Transparent
      Caption         =   "didici"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1335
      Left            =   3645
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblAudivi 
      BackStyle       =   0  'Transparent
      Caption         =   "Audivi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      Height          =   2115
      Left            =   120
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SplashMaximum As Integer
Private SplashValue As Integer

Private Sub Form_Load()
    SplashMaximum = 100
    SplashValue = 0
End Sub

Sub SetMaximumAndValue(myMaximum As Integer, myValue As Integer)
    SplashMaximum = myMaximum
    SplashValue = myValue
    UpdateControls
End Sub

Sub SetMaximum(myMaximum As Integer)
    SplashMaximum = myMaximum
    UpdateControls
End Sub

Sub SetValue(myValue As Integer)
    SplashValue = myValue
    UpdateControls
End Sub

Sub UpdateControls()
    On Error Resume Next
    
    Dim myLeft As Integer
    
    If SplashValue > SplashMaximum Then
        SplashValue = SplashMaximum
    End If
    
    myLeft = (7000 * (SplashValue / SplashMaximum)) + 360
    
    If myLeft >= 6615 Then
        shpHide.Visible = False
        lblAudivi.ForeColor = &H80C0FF
        lblVidi.ForeColor = &H80C0FF
        lblDidici.ForeColor = &H80FF&
        lblDidici.ZOrder vbBringToFront
    Else
        shpHide.Left = myLeft
        shpHide.Width = 6615 - myLeft
        
        If myLeft <= 2880 Then
            lblAudivi.ForeColor = &H80FF&
            lblVidi.ForeColor = &H80C0FF
            lblDidici.ForeColor = &H80C0FF
            lblAudivi.ZOrder vbBringToFront
        ElseIf myLeft <= 4560 Then
            lblAudivi.ForeColor = &H80C0FF
            lblVidi.ForeColor = &H80FF&
            lblDidici.ForeColor = &H80C0FF
            lblVidi.ZOrder vbBringToFront
        Else
            lblAudivi.ForeColor = &H80C0FF
            lblVidi.ForeColor = &H80C0FF
            lblDidici.ForeColor = &H80FF&
            lblDidici.ZOrder vbBringToFront
        End If
    End If
    
    lblAudiviText.ForeColor = lblAudivi.ForeColor
    lblVidiText.ForeColor = lblVidi.ForeColor
    lblDidiciText.ForeColor = lblDidici.ForeColor
End Sub
