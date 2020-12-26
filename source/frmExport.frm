VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmExport 
   Caption         =   "Audivididici"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox RTF 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmExport.frx":08CA
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Const WM_PASTE = &H302

Private Sub Form_Resize()
    RTF.Top = 120
    RTF.Left = 120
    RTF.Width = Me.Width - 360
    RTF.Height = Me.Height - 840
End Sub

Public Function LoadFile(FilePath As String, FileName As String, Pictures As Boolean) As Boolean
    Dim MyFreeFile As Integer
    Dim MyLine As String
    Dim MyWordNumber As Integer
    Dim TempMax As Integer
    Dim RetCode As Long
    Dim TempPicture As StdPicture
    
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
    
    Load frmSplash
    frmSplash.SetMaximumAndValue 1, 0
    frmSplash.Show

    If Not (LCase$(Right$(FileName, 4)) = ".avd" And FilePath <> "") Then
        MsgBox "Je hebt geen geldig bestand geselecteerd!", vbCritical + vbOKOnly, "Audivididici Creator"
        Unload frmSplash
        Unload Me
        Exit Function
    End If
    
    DelDir AppPath & "AVDTemp\"
    Dir AppPath
    
    '-- Init  Message Variables
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
    uZipFileName = FilePath & FileName
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
        
        RTF.TextRTF = ""
        RTF.SelStart = Len(RTF.Text)
        
        Do While Not EOF(MyFreeFile)
            Line Input #MyFreeFile, MyLine
            MyLine = Trim$(MyLine)
            RTF.SelRTF = Replace(Replace(MyLine, "\par ", " "), "\par}", "}")
            RTF.SelStart = Len(RTF.Text)
            RTF.SelText = " = "
            RTF.SelStart = Len(RTF.Text)
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            
            Line Input #MyFreeFile, MyLine
            RTF.SelRTF = Replace(Replace(MyLine, "\par ", " "), "\par}", "}")
            RTF.SelStart = Len(RTF.Text)
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            
            Line Input #MyFreeFile, MyLine
            
            If Pictures Then
                MyLine = Trim$(MyLine)
                If FExists(AppPath & "AVDTemp\" & MyLine) Then
                    Set TempPicture = StdFunctions.LoadPicture(AppPath & "AVDTemp\" & MyLine)
                    Clipboard.Clear
                    Clipboard.SetData TempPicture
                    RTF.SelText = " = "
                    RTF.SelStart = Len(RTF.Text)
                    SendMessage RTF.hwnd, WM_PASTE, 0, 0
                    RTF.SelStart = Len(RTF.Text)
                End If
            End If
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            Line Input #MyFreeFile, MyLine
            MyLine = Trim$(MyLine)
            
            RTF.SelText = vbCrLf
            RTF.SelStart = Len(RTF.Text)
            
            
            If EOF(MyFreeFile) Then GoTo DoEinde
            Line Input #MyFreeFile, MyLine

DoEinde:
            MyWordNumber = MyWordNumber + 1
            
            frmSplash.SetValue MyWordNumber + 1
            DoEvents
        Loop
    Close #MyFreeFile
    
    Unload frmSplash
    
    DelDir AppPath & "AVDTemp\"
    Dir AppPath
End Function
