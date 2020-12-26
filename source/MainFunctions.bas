Attribute VB_Name = "MainFunctions"
Option Explicit

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" _
    (ByVal lpFileName As String) As Long
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, _
    ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, _
    ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, _
    phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As _
    Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, _
    ByVal cbData As Long) As Long

Const KEY_WRITE = &H20006  '((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or
                           ' KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4

Const KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))
Const REG_OPENED_EXISTING_KEY = &H2

Const SHCNE_ASSOCCHANGED = &H8000000
Const SHCNF_IDLIST = 0

Public Const Demo = False
    
Public AppPath As String
Public LearnLanguages(255) As String
Public CurrentLearnLanguage As String
Public LearndataPath As String
Public CurrentLearnLanguageFile As String
Public CurrentLearnLanguagePath As String

Public Function FExists(OrigFile As String)
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    FExists = fs.FileExists(OrigFile)
End Function

Public Function DirExists(OrigFile As String)
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    DirExists = fs.FolderExists(OrigFile)
End Function

Sub Main()
    If App.PrevInstance Then
        MsgBox "Audivididici is al geopend!", vbCritical + vbOKOnly, "Audivididici"
        End
    End If
    
    AppPath = App.Path
    If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
        
        
    Load frmSplash
    frmSplash.SetMaximumAndValue 2, 1
    frmSplash.Show
    
    
    Load frmWords
    frmSplash.SetValue 2
    
    Load frmMain
    frmMain.InitControls Replace(Command, Chr$(34), "")
    frmMain.Show
    
    Unload frmSplash
    
End Sub

Public Sub DelDir(Path As String)
    On Error Resume Next
    
    Dim tempor As String
    
    If DirExists(Path) Then
        tempor = Dir(Path + "*.*", 22)
        Do While tempor <> ""
            If tempor <> "." And tempor <> ".." Then
                If (GetAttr(Path & tempor) And 16) = 0 Then
                    SetAttr Path & tempor, 0
                    DeleteFile Path & tempor
                Else
                    DelDir (Path & tempor & "\")
                    tempor = Dir(Path & "*.*", 22)
                End If
            End If
            tempor = Dir
        Loop
        RmDir Path
    End If
End Sub

' Create the new file association
'
' Extension is the extension to be registered (eg ".cad"
' ClassName is the name of the associated class (eg "CADDoc")
' Description is the textual description (eg "CAD Document"
' ExeProgram is the app that manages that extension (eg "c:\Cad\MyCad.exe")
'
' NOTE: requires CreateRegistryKey and SetRegistryValue functions

Sub CreateFileAssociation(ByVal Extension As String, ByVal ClassName As String, _
    ByVal Description As String, ByVal ExeProgram As String)
    Const HKEY_CLASSES_ROOT = &H80000000
    
    ' ensure that there is a leading dot
    If Left(Extension, 1) <> "." Then
        Extension = "." & Extension
    End If
   
    ' create a new registry key under HKEY_CLASSES_ROOT
    CreateRegistryKey HKEY_CLASSES_ROOT, Extension
    ' create a value for this key that contains the classname
    SetRegistryValue HKEY_CLASSES_ROOT, Extension, "", ClassName
    ' create a new key for the Class name
    CreateRegistryKey HKEY_CLASSES_ROOT, ClassName & "\Shell\Open\Command"
    ' set its value to the command line
    SetRegistryValue HKEY_CLASSES_ROOT, ClassName & "\Shell\Open\Command", "", _
        ExeProgram & " ""%1"""

    ' notify Windows that file associations have changed
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub

' Create a registry key, then close it
' Returns True if the key already existed, False if it was created

Function CreateRegistryKey(ByVal hKey As Long, ByVal KeyName As String) As _
    Boolean
    Dim handle As Long, disposition As Long
    
    If RegCreateKeyEx(hKey, KeyName, 0, 0, 0, 0, 0, handle, disposition) Then
        Err.Raise 1001, , "Unable to create the registry key"
    Else
        ' Return True if the key already existed.
        CreateRegistryKey = (disposition = REG_OPENED_EXISTING_KEY)
        ' Close the key.
        RegCloseKey handle
    End If
End Function

' Write or Create a Registry value
' returns True if successful
'
' Use KeyName = "" for the default value
'
' Value can be an integer value (REG_DWORD), a string (REG_SZ)
' or an array of binary (REG_BINARY). Raises an error otherwise.

Function SetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, value As Variant) As Boolean
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim Length As Long
    Dim retVal As Long
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then
        Exit Function
    End If

    ' three cases, according to the data type in Value
    Select Case VarType(value)
        Case vbInteger, vbLong
            lngValue = value
            retVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
        Case vbString
            strValue = value
            retVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
                Len(strValue))
        Case vbArray + vbByte
            binValue = value
            Length = UBound(binValue) - LBound(binValue) + 1
            retVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, _
                binValue(LBound(binValue)), Length)
        Case Else
            RegCloseKey handle
            Err.Raise 1001, , "Unsupported value type"
    End Select
    
    ' Close the key and signal success
    RegCloseKey handle
    ' signal success if the value was written correctly
    SetRegistryValue = (retVal = 0)
End Function
