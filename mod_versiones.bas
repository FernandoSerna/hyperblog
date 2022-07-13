Attribute VB_Name = "mod_versiones"
Option Explicit


Public Type FILEVERSIONINFO       'One invented by me
  path              As String
  Filename          As String
  Filesize          As String
  OSType            As String
  BinState          As String
  FileCreated       As String
  FileLastWritten   As String
  FileLastRead      As String
  CompanyName       As String
  FileDescription   As String
  FileVersion       As String     'The Binary File Version
  InternalName      As String
  LegalCopyright    As String
  OriginalFileName  As String
  ProductName       As String
  ProductVersion    As String     'The INFO Product Version
  Attributes        As String
End Type

'MICROSOFT STRUCTURES
Private Const OF_READ As Integer = &H0
Private Const OF_SHARE_DENY_NONE As Integer = &H40
Private Const OFS_MAXPATHNAME As Integer = 128

Private Type OFSTRUCTREC
  cBytes As Byte
  fFixedDisk As Byte
  nErrCode As Integer
  Reserved1 As Integer
  Reserved2 As Integer
  szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Type FILETIMEREC
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type SYSTEMTIMEREC
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
  dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
  dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
  dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
  dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
  dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
  dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
  dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
  dwFileType As Long             '  e.g. VFT_DRIVER
  dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
  dwFileDateMS As Long           '  e.g. 0
  dwFileDateLS As Long           '  e.g. 0
End Type

'Structure for Recycle Bin
Public Type SHFILEOPSTRUCT
  hwnd      As Long
  wFunc     As Long
  pFrom     As String
  pTo       As String
  fFlags    As Integer
  fAborted  As Boolean
  hNameMaps As Long
  sProgress As String
End Type

'Constants for Recycle Bin
Public Const FO_DELETE As Long = &H3
Public Const FOF_NOCONFIRMATION As Long = &H10
Public Const FOF_ALLOWUNDO As Long = &H40

'Operating system
'Public Const VOS__BASE = &H0&
'Public Const VOS_UNKNOWN = &H0&
Public Const VOS__WINDOWS16 = &H1&
Public Const VOS__PM16 = &H2&
Public Const VOS__PM32 = &H3&
Public Const VOS__WINDOWS32 = &H4&
Public Const VOS_DOS = &H10000
Public Const VOS_DOS_WINDOWS16 = &H10001
Public Const VOS_DOS_WINDOWS32 = &H10004
Public Const VOS_OS216 = &H20000
Public Const VOS_OS216_PM16 = &H20002
Public Const VOS_OS232 = &H30000
Public Const VOS_OS232_PM32 = &H30003
Public Const VOS_NT = &H40000
Public Const VOS_NT_WINDOWS32 = &H40004

'File Attributes
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20        'An archive file (which most files are).':( As Integer ?
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800    'A file residing in a compressed drive or directory.':( As Integer ?
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10      'A directory instead of a file.':( As Integer ?
Public Const FILE_ATTRIBUTE_HIDDEN = &H2          'A hidden file, not normally visible to the user.':( As Integer ?
Public Const FILE_ATTRIBUTE_NORMAL = &H80         'An attribute-less file (cannot be combined with other attributes).':( As Integer ?
Public Const FILE_ATTRIBUTE_READONLY = &H1        'A read-only file.':( As Integer ?
Public Const FILE_ATTRIBUTE_SYSTEM = &H4          'A system file, used exclusively by the operating system.':( As Integer ?

'FileState
Public Const VS_FF_DEBUG = &H1&
Public Const VS_FF_PRERELEASE = &H2&
Public Const VS_FF_PATCHED = &H4&
Public Const VS_FF_PRIVATEBUILD = &H8&
Public Const VS_FF_INFOINFERRED = &H10&
Public Const VS_FF_SPECIALBUILD = &H20&
Public Const VS_FFI_FILEFLAGSMASK = &H3F&
 
'KERNEL32.DLL FUNCTIONS
    
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (dest As Any, ByVal Source As Long, ByVal Length As Long)

Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
        (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" _
                (lpFileTime As FILETIMEREC, lpSystemTime As SYSTEMTIMEREC) As Long

Private Declare Function GetFileTime Lib "kernel32" _
                (ByVal hFile As Long, lpCreationTime As FILETIMEREC, _
                lpLastAccessTime As FILETIMEREC, lpLastWriteTime As FILETIMEREC) As Long

Private Declare Function OpenFile Lib "kernel32" _
                (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCTREC, ByVal wStyle As Long) As Long

Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
              (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Private Declare Function lclose Lib "kernel32" Alias "_lclose" _
                (ByVal hFile As Long) As Long
  
'VERSION.DLL FUNCTIONS
Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
        (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long

Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" _
        (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" _
        (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

'SHELL32.DLL FUNCTIONS
  
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
              (lpFileOp As SHFILEOPSTRUCT) As Long

'=========================================================================================================
Private Function hex2(v As Byte) As String

  hex2 = Right$("00" & Hex$(v), 2)

End Function

Private Function FmtFileTime(ft As FILETIMEREC) As String

 Dim st As SYSTEMTIMEREC
    
  If FileTimeToSystemTime(ft, st) <> 0 Then
'    FmtFileTime = Format$(st.wMonth, "00") _
'                  & "/" _
'                  & Format$(st.wDay, "00") _
'                  & "/" _
'                  & Format$(st.wYear, "0000") _    'Mark! for your US dates but they break the sort

    FmtFileTime = Format$(st.wYear, "0000") _
                  & "/" _
                  & Format$(st.wMonth, "00") _
                  & "/" _
                  & Format$(st.wDay, "00") _
                  & " " _
                  & Format$(st.wHour, "00") _
                  & ":" _
                  & Format$(st.wMinute, "00") _
                  & ":" _
                  & Format$(st.wSecond, "00") _
                  & "." _
                  & Format$(st.wMilliseconds, "000")
   Else
    FmtFileTime = "?"
  End If

End Function

Private Function GetFileTimeStamps(FVI As FILEVERSIONINFO) As Integer

 Dim hFile As Integer
 Dim FileStruct As OFSTRUCTREC
 Dim CreationTime As FILETIMEREC
 Dim LastAccessTime As FILETIMEREC
 Dim LastWriteTime As FILETIMEREC
 Dim rc As Integer
  
  ' Open it to get a stream handle
  hFile = OpenFile(FVI.path, FileStruct, OF_READ Or OF_SHARE_DENY_NONE)
  If hFile <> 0 Then
    If GetFileTime(hFile, CreationTime, LastAccessTime, LastWriteTime) Then
      FVI.FileCreated = FmtFileTime(CreationTime)
      FVI.FileLastRead = FmtFileTime(LastAccessTime)
      FVI.FileLastWritten = FmtFileTime(LastWriteTime)
    End If
    rc = lclose(hFile)
   Else
    rc = 0
  End If
  GetFileTimeStamps = rc

End Function

Private Function FmtOSFlags(OSFlags As Long) As String

  If OSFlags = 0 Then
    FmtOSFlags = "?"
   Else
    Select Case OSFlags
     Case VOS__WINDOWS16
      FmtOSFlags = "WIN16"
     Case VOS__PM16
      FmtOSFlags = "PM16"
     Case VOS__PM32
      FmtOSFlags = "PM32"
     Case VOS__WINDOWS32
      FmtOSFlags = "WIN32"
     Case VOS_DOS
      FmtOSFlags = "DOS"
     Case VOS_DOS_WINDOWS16
      FmtOSFlags = "DOS16"
     Case VOS_DOS_WINDOWS32
      FmtOSFlags = "DOS32"
     Case VOS_OS216
      FmtOSFlags = "OS216"
     Case VOS_OS216_PM16
      FmtOSFlags = "OS2PM16"
     Case VOS_OS232
      FmtOSFlags = "OS232"
     Case VOS_OS232_PM32
      FmtOSFlags = "OS2PM32"
     Case VOS_NT
      FmtOSFlags = "NT"
     Case VOS_NT_WINDOWS32
      FmtOSFlags = "NTW32"
     Case Else
      FmtOSFlags = "OTHER"
    End Select
  End If

End Function

Private Function FmtBinFlags(BinFlags As Long) As String

  BinFlags = BinFlags And VS_FFI_FILEFLAGSMASK
  FmtBinFlags = ""
  If BinFlags And VS_FF_DEBUG Then
    FmtBinFlags = FmtBinFlags & "D"
   ElseIf BinFlags And VS_FF_PRERELEASE = VS_FF_PRERELEASE Then
    FmtBinFlags = FmtBinFlags & "b"
   ElseIf BinFlags And VS_FF_PATCHED = VS_FF_PATCHED Then
    FmtBinFlags = FmtBinFlags & "p"
   ElseIf BinFlags And VS_FF_PRIVATEBUILD = VS_FF_PRIVATEBUILD Then
    FmtBinFlags = FmtBinFlags & "P"
   ElseIf BinFlags And VS_FF_INFOINFERRED = VS_FF_INFOINFERRED Then
    FmtBinFlags = FmtBinFlags & "I"
   ElseIf BinFlags And VS_FF_SPECIALBUILD = VS_FF_SPECIALBUILD Then
    FmtBinFlags = FmtBinFlags & "S"
  End If

End Function

Private Function FmtVersion(MSLong As Long, LSLong As Long) As String

  FmtVersion = Format$((MSLong And &HFFFF0000) \ 65536, "####0.") _
               & Format$((MSLong And &HFFFF&), "###00.") _
               & Format$((LSLong And &HFFFF0000) \ 65536, "###00.") _
               & Format$((LSLong And &HFFFF&), "#0000")

End Function

Private Function FmtAttributes(ByVal Attribs As Long) As String     'thanks Mark
    
  FmtAttributes = ""
    
  If Attribs And FILE_ATTRIBUTE_ARCHIVE Then
    FmtAttributes = "Archive"
  End If
        
  If Attribs And FILE_ATTRIBUTE_COMPRESSED Then
    FmtAttributes = FmtAttributes & ", Compressed"
  End If
    
  If Attribs And FILE_ATTRIBUTE_DIRECTORY Then
    FmtAttributes = FmtAttributes & ", Directory"
  End If
        
  If Attribs And FILE_ATTRIBUTE_HIDDEN Then
    FmtAttributes = FmtAttributes & ", Hidden"
  End If
    
  If Attribs And FILE_ATTRIBUTE_NORMAL Then
    FmtAttributes = FmtAttributes & ", Normal"
  End If
    
  If Attribs And FILE_ATTRIBUTE_READONLY Then
    FmtAttributes = FmtAttributes & ", Read-Only"
  End If
        
  If Attribs And FILE_ATTRIBUTE_SYSTEM Then
    FmtAttributes = FmtAttributes & ", System"
  End If
    
  If Left$(FmtAttributes, 2) = ", " Then FmtAttributes = Mid$(FmtAttributes, 3)

End Function

'We will get the complete version info into FILEVERSIONINFO
'a return code of <=0 implies we had a failure, >0 all OK
Public Function GetFVInfo(ByRef FVI As FILEVERSIONINFO) As Long

 Dim q As Long, i As Long, vptr As Long, vlen As Long, vsffi As VS_FIXEDFILEINFO
 Dim InfoSize As Long, Info() As Byte, wsp(0 To 15) As Byte, Buf As String
 Dim SubBlock As String, Lang_Charset As String, VersionInfo(0 To 7) As String
  
  FVI.Filesize = "?"
  FVI.BinState = "?"
  FVI.OSType = "?"
  FVI.CompanyName = ""
  FVI.FileDescription = ""
  FVI.FileVersion = "?.?.?.?"
  FVI.FileCreated = "?"
  FVI.InternalName = ""
  FVI.LegalCopyright = ""
  FVI.OriginalFileName = ""
  FVI.ProductName = ""
  FVI.ProductVersion = "?.?.?.?"
  FVI.Attributes = "?"
    
  On Error Resume Next                                      'get the filesize
   FVI.Filesize = Format$(FileLen(FVI.path), "###,###,###")
   FVI.Attributes = FmtAttributes(GetAttr(FVI.path))         'get file attributes
  On Error GoTo 0
  
  Call GetFileTimeStamps(FVI)                            'get file timestamps :-)
  'now get the other version information
  InfoSize = GetFileVersionInfoSize(FVI.path, q)
If InfoSize <= 0 Then
    GetFVInfo = -1      'version not available
    Exit Function
  End If
  
  ReDim Info(0 To InfoSize) As Byte
  If GetFileVersionInfo(FVI.path, q, InfoSize, Info(0)) <= 0 Then
    GetFVInfo = -2      'read versioninfo failed
    Exit Function
  End If
  
  SubBlock = "\"    'Root for FixedFileInfo
  If VerQueryValue(Info(0), SubBlock, vptr, vlen) > 0 Then
    If vlen > 0 Then
      Call CopyMemory(vsffi, vptr, vlen)
      FVI.FileVersion = FmtVersion(vsffi.dwFileVersionMS, vsffi.dwFileVersionLS)
      FVI.ProductVersion = FmtVersion(vsffi.dwProductVersionMS, vsffi.dwProductVersionLS)
      FVI.OSType = FmtOSFlags(vsffi.dwFileOS)
      FVI.BinState = FmtBinFlags(vsffi.dwFileFlags)
    End If
  End If
  
  SubBlock = "\VarFileInfo\Translation"
  If VerQueryValue(Info(0), SubBlock, vptr, vlen) <= 0 Then
    Lang_Charset = "040904E4"       'read translation key failed so assume MS default
   Else
    'vptr is a pointer to four 4 bytes of Hex number,
    'first two bytes are language id, and last two bytes are code page.
    'However, VerQueryValue needs a  string of 4 hex digits,
    ' the first two characters correspond to the language id and
    ' the last two characters correspond to the code page id.
    'now we change the order of the language id and code page
    'and convert it into a string representation.
    'For example, it may look like 040904E4
    'Or to pull it all apart:
    '04------        = SUBLANG_ENGLISH_USA
    '--09----        = LANG_ENGLISH
    ' ----04E4 = 1252 = Codepage for Windows:Multilingual
    Call CopyMemory(wsp(0), vptr, vlen)
    Lang_Charset = hex2(wsp(1)) & hex2(wsp(0)) & hex2(wsp(3)) & hex2(wsp(2))
  End If
  
  VersionInfo(0) = "CompanyName"
  VersionInfo(1) = "FileDescription"
  VersionInfo(2) = "FileVersion"
  VersionInfo(3) = "InternalName"
  VersionInfo(4) = "LegalCopyright"
  VersionInfo(5) = "OriginalFileName"
  VersionInfo(6) = "ProductName"
  VersionInfo(7) = "ProductVersion"

  For i = 0 To 7
    Buf = String$(255, 0)
    SubBlock = "\StringFileInfo\" & Lang_Charset & "\" & VersionInfo(i)
    If VerQueryValue(Info(0), SubBlock, vptr, vlen) = 0 Then
      GetFVInfo = -3      'read subblock key failed
      Exit Function
    End If
    If vlen > 0 Then
      Call lstrcpy(Buf, vptr)
      VersionInfo(i) = Mid$(Buf, 1, InStr(Buf, Chr$(0)) - 1)
     Else
      VersionInfo(i) = ""
    End If
  Next i
  
  'fill the outbound array  - we dont actually use all of these in the main program
  FVI.CompanyName = VersionInfo(0)
  FVI.FileDescription = VersionInfo(1)
  If FVI.FileVersion = "?.?.?.?" Then
    FVI.FileVersion = VersionInfo(2)
  End If
    
  FVI.InternalName = VersionInfo(3)
  FVI.LegalCopyright = VersionInfo(4)
  FVI.OriginalFileName = VersionInfo(5)
  FVI.ProductName = VersionInfo(6)
  If FVI.ProductVersion < VersionInfo(7) Then
    FVI.ProductVersion = VersionInfo(7)
  End If
  GetFVInfo = 1

End Function

'Returns True if successful     'moves a file to the Recycle Bin    Thanks Mark
Public Function Erase2RecycleBin(fileSpec As String) As Boolean

 Dim SHFileOp As SHFILEOPSTRUCT
    
  With SHFileOp
    .wFunc = FO_DELETE
    .pFrom = fileSpec & vbNullChar
    .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
  End With
    
  Erase2RecycleBin = (SHFileOperation(SHFileOp) = 0)
  
End Function

Sub proversion()

End Sub
