Attribute VB_Name = "Conversion"
Option Explicit

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" _
    (ByVal lpFileName As String) As Long


Public Sub ConvertDir(Path As String, Compression As Integer)
    On Error Resume Next
    
    Dim tempor As String
    Dim List(9999) As String
    Dim i As Integer
    Dim Length As Integer
    
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    
    i = 0
    
    If DirExists(Path) Then
        tempor = Dir(Path & "*.*", 22)
        Do While tempor <> ""
            If tempor <> "." And tempor <> ".." Then
                If (GetAttr(Path & tempor) And 16) = 0 Then
                    List(i) = tempor
                    i = i + 1
                End If
            End If
            tempor = Dir
            DoEvents
        Loop
    End If
    
    Length = i
    
    For i = 0 To Length - 1
        ConvertImport Path, List(i), Compression, True
    Next i
    
End Sub



Public Function ConvertImport(FilePath As String, FileName As String, Compression As Integer, Optional Discrete As Boolean = False) As String
    ConvertImport = ""
    If LCase$(Right$(FileName, 4)) = ".oh4" Or LCase$(Right$(FileName, 4)) = ".ohw" Then
        ConvertImport = ConvertOH4(FilePath, FileName, Compression, Discrete)
    ElseIf LCase$(Right$(FileName, 4)) = ".t2k" Then
        ConvertImport = ConvertT2K(FilePath, FileName, Compression, Discrete)
    End If
End Function


Function ConvertOH4(FilePath As String, FileName As String, Compression As Integer, Optional Discrete As Boolean = False) As String
    Dim MyFreeFileInput As Integer
    Dim MyFreeFileOutput As Integer
    Dim MyLine As String
    Dim Seperator As Integer
    Dim OutputFile As String
    Dim Questions(999) As String
    Dim Answers(999) As String
    Dim i As Integer
    Dim Length As Integer
    Dim RetCode As Integer
    Dim Language As String
    
    MyFreeFileInput = FreeFile()
    MyFreeFileOutput = FreeFile()
    
    If Right$(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    
    If Not FExists(FilePath & FileName) Then
        If Not Discrete Then MsgBox "Bestand niet gevonden!", vbCritical, "Audivididici Creator"
        ConvertOH4 = ""
        Exit Function
    End If

    If Not ((LCase$(Right$(FileName, 4)) = ".oh4" Or LCase$(Right$(FileName, 4)) = ".ohw") And FileName <> "") Then
        If Not Discrete Then MsgBox "Je hebt geen geldig bestand geselecteerd!", vbCritical + vbOKOnly, "Audivididici Creator"
        ConvertOH4 = ""
        Exit Function
    End If
    
    OutputFile = Left$(FileName, Len(FileName) - 4) & ".avd"
    
    If FExists(FilePath & OutputFile) Then
        If Not Discrete Then
            If MsgBox(OutputFile & " bestaat al! Weet je zeker dat je wil overschrijven?", vbQuestion + vbYesNo, "Audivididici") = vbNo Then
                ConvertOH4 = ""
                Exit Function
            End If
        End If
        DeleteFile FilePath & OutputFile
    End If
    
    DelDir AppPath & "AVDTemp\"
    Dir AppPath
    
    Open FilePath & FileName For Input As #MyFreeFileInput
        If LCase$(Right$(FileName, 4)) = ".oh4" Then
            Line Input #MyFreeFileInput, MyLine
            MyLine = Trim$(MyLine)
            Seperator = InStr(1, MyLine, " = ")
            Questions(0) = Trim$(Left$(MyLine, Seperator))
            Answers(0) = Trim$(Right$(MyLine, Len(MyLine) - Seperator - 2))
            
            If InStr(1, LCase(Questions(0)), "greek") > 0 Or InStr(1, LCase(Questions(0)), "symbol") > 0 Then
                Language = "G"
            Else
                Language = "L"
            End If
            
            If InStr(1, LCase(Answers(0)), "greek") > 0 Or InStr(1, LCase(Answers(0)), "symbol") > 0 Then
                Language = Language & "G"
            Else
                Language = Language & "L"
            End If
            
            Questions(0) = ""
            Answers(0) = ""
        Else
            Language = "LL"
        End If
        
        i = 0
        
        Do While Not EOF(MyFreeFileInput)
            Line Input #MyFreeFileInput, MyLine
            MyLine = Trim$(MyLine)
            Seperator = InStr(1, MyLine, " = ")
            Questions(i) = Trim$(Left$(MyLine, Seperator))
            Answers(i) = Trim$(Right$(MyLine, Len(MyLine) - Seperator - 2))
            
            If Mid(Language, 1, 1) = "G" Then
                Questions(i) = LatinToGreek(Questions(i))
            End If
            Questions(i) = MakeUnicodeRTF(Questions(i))
            
            If Mid(Language, 2, 1) = "G" Then
                Answers(i) = LatinToGreek(Answers(i))
            End If
            Answers(i) = MakeUnicodeRTF(Answers(i))
            
            If Questions(i) <> "" Or Answers(i) <> "" Then
                i = i + 1
            End If
        Loop
    Close #MyFreeFileInput
    
    Length = i
    
    If Not DirExists(AppPath & "AVDTemp") Then
        MkDir AppPath & "AVDTemp"
    End If
    
    Open AppPath & "AVDTemp\info.txt" For Output As #MyFreeFileOutput
        Print #MyFreeFileOutput, Trim$(Str$(Length))
        
        For i = 0 To Length - 1
            Print #MyFreeFileOutput, "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Gentium;}}\viewkind4\pard\f0\lang1032\b\fs38\qc " & Questions(i) & "\par}"
            Print #MyFreeFileOutput, "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Gentium;}}\viewkind4\uc1\pard\f0\lang1032\fs32\qc " & Answers(i) & "\par}"
            Print #MyFreeFileOutput, ""
            Print #MyFreeFileOutput, ""
            Print #MyFreeFileOutput, Language
        Next i
    Close #MyFreeFileOutput
    
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
    zZipFileName = FilePath & OutputFile
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
    
    ConvertOH4 = FilePath & OutputFile
End Function
    



Function ConvertT2K(FilePath As String, FileName As String, Compression As Integer, Optional Discrete As Boolean = False) As String
    Dim MyFreeFileInput As Integer
    Dim MyFreeFileOutput As Integer
    Dim MyLine As String
    Dim MyString As String
    Dim Position As Integer
    Dim Position2 As Integer
    Dim PositionNext As Integer
    Dim PositionPrevious As Integer
    Dim OutputFile As String
    Dim i As Integer
    Dim MyName As String
    Dim RetCode As Integer
    Dim Length As Integer
    Dim Language As String
    Dim Questions(999) As String
    Dim Answers(999) As String
    Dim Pictures(999) As String
    Dim Sounds(999) As String
    
    MyFreeFileInput = FreeFile()
    MyFreeFileOutput = FreeFile()
    
    If Right$(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    
    If Not FExists(FilePath & FileName) Then
        If Not Discrete Then MsgBox "Bestand niet gevonden!", vbCritical, "Audivididici Creator"
        ConvertT2K = ""
        Exit Function
    End If

    If Not (LCase$(Right$(FileName, 4)) = ".t2k" And FileName <> "") Then
        If Not Discrete Then MsgBox "Je hebt geen geldig bestand geselecteerd!", vbCritical + vbOKOnly, "Audivididici Creator"
        ConvertT2K = ""
        Exit Function
    End If
    
    OutputFile = Left$(FileName, Len(FileName) - 4) & ".avd"
    
    If FExists(FilePath & OutputFile) Then
        If Not Discrete Then
            If MsgBox(OutputFile & " bestaat al! Weet je zeker dat je wil overschrijven?", vbQuestion + vbYesNo, "Audivididici") = vbNo Then
                ConvertT2K = ""
                Exit Function
            End If
        End If
        DeleteFile FilePath & OutputFile
    End If
    
    DelDir AppPath & "AVDTemp\"
    Dir AppPath
    
    MyString = ""
    
    Open FilePath & FileName For Input As #MyFreeFileInput
        Do Until EOF(MyFreeFileInput)
            Line Input #MyFreeFileInput, MyLine
            MyString = MyString & MyLine
        Loop
    Close #MyFreeFileInput
    
    If Not DirExists(AppPath & "AVDTemp") Then
        MkDir AppPath & "AVDTemp"
    End If
    
    
    Position = InStrRev(MyString, "<font_question>") + Len("<font_question>")
    Position2 = InStr(Position, MyString, "</font_question>")
    MyLine = Mid$(MyString, Position, Position2 - Position)
    If InStr(1, LCase(MyLine), "greek") > 0 Or InStr(1, LCase(MyLine), "symbol") > 0 Then
        Language = "G"
    Else
        Language = "L"
    End If
    
    Position = InStrRev(MyString, "<font_answer>") + Len("<font_answer>")
    Position2 = InStr(Position, MyString, "</font_answer>")
    MyLine = Mid$(MyString, Position, Position2 - Position)
    If InStr(1, LCase(MyLine), "greek") > 0 Or InStr(1, LCase(MyLine), "symbol") > 0 Then
        Language = Language & "G"
    Else
        Language = Language & "L"
    End If
    
    PositionNext = InStr(1, MyString, "<question>") + Len("<question>")
    i = 0
    Do While PositionNext > Len("<question>")
        Position = PositionNext
        Position2 = InStr(Position, MyString, "</question>")
        MyLine = Mid$(MyString, Position, Position2 - Position)
        If Mid(Language, 1, 1) = "G" Then
            MyLine = LatinToGreek(MyLine)
        End If
        MyLine = MakeUnicodeRTF(MyLine)
        Questions(i) = "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Gentium;}}\viewkind4\pard\lang1032\f0\b\fs38\qc " & MyLine & "\par}"
        
        PositionNext = InStr(Position2, MyString, "<question>") + Len("<question>")
        PositionPrevious = Position2
        
        Position = InStr(PositionPrevious, MyString, "<answer>") + Len("<answer>")
        Position2 = InStr(Position, MyString, "</answer>")
        MyLine = Mid$(MyString, Position, Position2 - Position)
        If Mid(Language, 2, 1) = "G" Then
            MyLine = LatinToGreek(MyLine)
        End If
        MyLine = MakeUnicodeRTF(MyLine)
        Answers(i) = "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Gentium;}}\viewkind4\uc1\pard\f0\lang1032\fs32\qc " & MyLine & "\par}"
        
        Pictures(i) = ""
        
        Position = InStr(PositionPrevious, MyString, "<picturefile>") + Len("<picturefile>")
        If Position > Len("<picturefile>") And (Position < PositionNext Or PositionNext <= 0) Then
            Position2 = InStr(Position, MyString, "</picturefile>")
            MyLine = Mid$(MyString, Position, Position2 - Position)
            If Right(MyLine, 4) = ".jpg" Or Right(MyLine, 5) = ".jpeg" Or Right(MyLine, 4) = ".bmp" Or Right(MyLine, 4) = ".gif" Then
                If FExists(MyLine) Then
                    MyName = Right(MyLine, Len(MyLine) - InStrRev(MyLine, "\"))
                    FileCopy MyLine, AppPath & "AVDTemp\" & MyName
                    Pictures(i) = MyName
                End If
            End If
        End If
        
        Sounds(i) = ""
        
        Position = InStr(PositionPrevious, MyString, "<soundfile>") + Len("<soundfile>")
        If Position > Len("<soundfile>") And (Position < PositionNext Or PositionNext <= 0) Then
            Position2 = InStr(Position, MyString, "</soundfile>")
            MyLine = Mid$(MyString, Position, Position2 - Position)
            If Right(MyLine, 4) = ".wav" Or Right(MyLine, 4) = ".mp3" Then
                If FExists(MyLine) Then
                    MyName = Right(MyLine, Len(MyLine) - InStrRev(MyLine, "\"))
                    FileCopy MyLine, AppPath & "AVDTemp\" & MyName
                    Sounds(i) = MyName
                End If
            End If
        End If
        i = i + 1
    Loop
    
    Length = i
    
    Open AppPath & "AVDTemp\info.txt" For Output As #MyFreeFileOutput
        Print #MyFreeFileOutput, Trim$(Str$(Length))
        
        For i = 0 To Length - 1
            Print #MyFreeFileOutput, Questions(i)
            Print #MyFreeFileOutput, Answers(i)
            Print #MyFreeFileOutput, Pictures(i)
            Print #MyFreeFileOutput, Sounds(i)
            Print #MyFreeFileOutput, Language
        Next i
    Close #MyFreeFileOutput
    
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
    zZipFileName = FilePath & OutputFile
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
    
    ConvertT2K = FilePath & OutputFile
End Function

Function LatinToGreek(RTFLatin As String) As String
    LatinToGreek = RTFLatin
    LatinToGreek = Replace(LatinToGreek, "u", "\u965?") 'upsilon
    LatinToGreek = Replace(LatinToGreek, "a", "\u945?") 'alpha
    LatinToGreek = Replace(LatinToGreek, "b", "\u946?") 'beta
    LatinToGreek = Replace(LatinToGreek, "c", "\u967?") 'chi
    LatinToGreek = Replace(LatinToGreek, "d", "\u948?") 'delta
    LatinToGreek = Replace(LatinToGreek, "e", "\u949?") 'epsilon
    LatinToGreek = Replace(LatinToGreek, "f", "\u966?") 'phi
    LatinToGreek = Replace(LatinToGreek, "g", "\u947?") 'gamma
    LatinToGreek = Replace(LatinToGreek, "h", "\u951?") 'eta
    LatinToGreek = Replace(LatinToGreek, "i", "\u953?") 'iota
    LatinToGreek = Replace(LatinToGreek, "j", "\u966?") 'phi
    LatinToGreek = Replace(LatinToGreek, "k", "\u954?") 'kappa
    LatinToGreek = Replace(LatinToGreek, "l", "\u955?") 'labda
    LatinToGreek = Replace(LatinToGreek, "m", "\u956?") 'mu
    LatinToGreek = Replace(LatinToGreek, "n", "\u957?") 'nu
    LatinToGreek = Replace(LatinToGreek, "o", "\u959?") 'omikron
    LatinToGreek = Replace(LatinToGreek, "p", "\u960?") 'pi
    LatinToGreek = Replace(LatinToGreek, "q", "\u952?") 'theta
    LatinToGreek = Replace(LatinToGreek, "r", "\u961?") 'rho
    LatinToGreek = Replace(LatinToGreek, "s", "\u963?") 'sigma
    LatinToGreek = Replace(LatinToGreek, "t", "\u964?") 'tau
    LatinToGreek = Replace(LatinToGreek, "v", "\u962?") 'sigma eind
    LatinToGreek = Replace(LatinToGreek, "w", "\u969?") 'omega
    LatinToGreek = Replace(LatinToGreek, "x", "\u958?") 'xi
    LatinToGreek = Replace(LatinToGreek, "y", "\u968?") 'psi
    LatinToGreek = Replace(LatinToGreek, "z", "\u950?") 'zeta
    LatinToGreek = Replace(LatinToGreek, "A", "\u913?") 'alpha
    LatinToGreek = Replace(LatinToGreek, "B", "\u914?") 'beta
    LatinToGreek = Replace(LatinToGreek, "C", "\u935?") 'chi
    LatinToGreek = Replace(LatinToGreek, "D", "\u916?") 'delta
    LatinToGreek = Replace(LatinToGreek, "E", "\u917?") 'epsilon
    LatinToGreek = Replace(LatinToGreek, "F", "\u934?") 'phi
    LatinToGreek = Replace(LatinToGreek, "G", "\u915?") 'gamma
    LatinToGreek = Replace(LatinToGreek, "H", "\u919?") 'eta
    LatinToGreek = Replace(LatinToGreek, "I", "\u921?") 'iota
    LatinToGreek = Replace(LatinToGreek, "J", "\u934?") 'phi
    LatinToGreek = Replace(LatinToGreek, "K", "\u922?") 'kappa
    LatinToGreek = Replace(LatinToGreek, "L", "\u923?") 'labda
    LatinToGreek = Replace(LatinToGreek, "M", "\u924?") 'mu
    LatinToGreek = Replace(LatinToGreek, "N", "\u925?") 'nu
    LatinToGreek = Replace(LatinToGreek, "O", "\u927?") 'omikron
    LatinToGreek = Replace(LatinToGreek, "P", "\u928?") 'pi
    LatinToGreek = Replace(LatinToGreek, "Q", "\u920?") 'theta
    LatinToGreek = Replace(LatinToGreek, "R", "\u929?") 'rho
    LatinToGreek = Replace(LatinToGreek, "S", "\u931?") 'sigma
    LatinToGreek = Replace(LatinToGreek, "T", "\u932?") 'tau
    LatinToGreek = Replace(LatinToGreek, "U", "\u933?") 'upsilon
    LatinToGreek = Replace(LatinToGreek, "V", "\u962?") 'sigma eind
    LatinToGreek = Replace(LatinToGreek, "W", "\u937?") 'omega
    LatinToGreek = Replace(LatinToGreek, "X", "\u926?") 'xi
    LatinToGreek = Replace(LatinToGreek, "Y", "\u936?") 'psi
    LatinToGreek = Replace(LatinToGreek, "Z", "\u918?") 'zeta
End Function

Function MakeUnicodeRTF(myRTF As String) As String
    Dim i As Integer
    MakeUnicodeRTF = ""
    For i = 1 To Len(myRTF)
        If (AscW(Mid(myRTF, i, 1)) < 128) Then
            MakeUnicodeRTF = MakeUnicodeRTF & Mid(myRTF, i, 1)
        Else
            MakeUnicodeRTF = MakeUnicodeRTF & "\u" & AscW(Mid(myRTF, i, 1)) & "?"
        End If
    Next i
End Function
