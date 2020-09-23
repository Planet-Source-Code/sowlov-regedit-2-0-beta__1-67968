VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistryAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
                                                         '
Option Explicit

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, ByVal lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ByVal ReturnLength As Long) As Long

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_INVALID_PARAMETER = 87&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234&
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_CREATE_LINK = &H20&
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const REG_FORCE_RESTORE = &H8
Private Const TOKEN_ADJUST_PRIVLEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const SE_RESTORE_NAME = "SeRestorePrivilege"
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(1) As LUID_AND_ATTRIBUTES
End Type

Public Enum rcMainKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Enum rcRegType
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
    REG_RESOURCE_LIST = 8
    REG_FULL_RESOURCE_DESCRIPTOR = 9
    REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum

Public Enum convertBINType
    BIN_Bin = 0
    BIN_Array = 1
    BIN_Dot = 2
    BIN_String = 3
End Enum

Private hKey             As Long
Private mainKey          As Long
Private sKey             As String
Private createNoExists   As Boolean
Private lReturn As Long

Private m_hToken As Long
Private m_TP As TOKEN_PRIVILEGES
Private m_RestoreLuid As LUID
Private m_BackupLuid As LUID

Private Const SITUATION_BASE As Long = 13800

Public Property Let CreateKeyIfDoesntExists(ByVal offon As Boolean)
Attribute CreateKeyIfDoesntExists.VB_Description = "Let/Get. Should key will be created if no exists (when writting data to registry)"
    createNoExists = offon
End Property

Public Property Get CreateKeyIfDoesntExists() As Boolean
    CreateKeyIfDoesntExists = createNoExists
End Property

Private Function GetKeys(sPath As String, sKey As String) As rcMainKey
Dim pos As Long, mk As String
    
    sPath = Replace$(sPath, "HKEY_CURRENT_USER", "HKCU", , , 1)
    sPath = Replace$(sPath, "HKEY_LOCAL_MACHINE", "HKLM", , , 1)
    sPath = Replace$(sPath, "HKEY_CLASSES_ROOT", "HKCR", , , 1)
    sPath = Replace$(sPath, "HKEY_USERS", "HKUS", , , 1)
    sPath = Replace$(sPath, "HKEY_PERFORMANCE_DATA", "HKPD", , , 1)
    sPath = Replace$(sPath, "HKEY_DYN_DATA", "HKDD", , , 1)
    sPath = Replace$(sPath, "HKEY_CURRENT_CONFIG", "HKCC", , , 1)
    
    pos = InStr(1, sPath, "\")

    If (pos = 0) Then
        mk = UCase$(sPath)
        sKey = ""
    Else
        mk = UCase$(Left$(sPath, 4))
        sKey = Right$(sPath, Len(sPath) - pos)
    End If
    
    Select Case mk
        Case "HKCU": GetKeys = HKEY_CURRENT_USER
        Case "HKLM": GetKeys = HKEY_LOCAL_MACHINE
        Case "HKCR": GetKeys = HKEY_CLASSES_ROOT
        Case "HKUS": GetKeys = HKEY_USERS
        Case "HKPD": GetKeys = HKEY_PERFORMANCE_DATA
        Case "HKDD": GetKeys = HKEY_DYN_DATA
        Case "HKCC": GetKeys = HKEY_CURRENT_CONFIG
    End Select
    
End Function

Public Function CreateKey(ByVal sPath As String) As Long
Attribute CreateKey.VB_Description = "Creates key in registry..."
    hKey = GetKeys(sPath, sKey)
    
    If (RegCreateKey(hKey, sKey, mainKey) = ERROR_SUCCESS) Then
        RegCloseKey mainKey
        CreateKey = mainKey
    Else
        CreateKey = 0
    End If
End Function

Public Function KillKey(ByVal sPath As String) As Long
Attribute KillKey.VB_Description = "Deletes key from registry..."
    Dim sKeys() As String, nKeys As Long, nIdx As Long
    nKeys = EnumKeys(sPath, sKeys)
    If nKeys > 0 Then
        For nIdx = 0 To nKeys - 1
            KillKey sPath & "\" & sKeys(nIdx)
        Next nIdx
    End If
    hKey = GetKeys(sPath, sKey)
    
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_ALL_ACCESS, mainKey) = ERROR_SUCCESS) Then
        RegDeleteKey mainKey, ""
        RegCloseKey mainKey
        KillKey = mainKey
    Else
        KillKey = 0
    End If
End Function

Public Function KeyExists(ByVal sPath As String) As Boolean
    hKey = GetKeys(sPath, sKey)
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_ALL_ACCESS, mainKey) = ERROR_SUCCESS) Then
        KeyExists = True
        RegCloseKey mainKey
    Else
        KeyExists = False
    End If
End Function

Public Function EnumKeys(ByVal sPath As String, ByRef Key() As String) As Long
Attribute EnumKeys.VB_Description = "Enumerates subkeys of some key..."
    Dim sName As String, retVal As Long
    
    hKey = GetKeys(sPath, sKey)
    
    Erase Key
    
    If (RegOpenKey(hKey, sKey, mainKey) = ERROR_SUCCESS) Then

        EnumKeys = 0
        sName = String(255, 0)
        retVal = Len(sName)
        
        lReturn = RegEnumKeyEx(mainKey, EnumKeys, sName, retVal, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&)
        While lReturn <> ERROR_BADDB And lReturn <> ERROR_BADKEY And lReturn <> ERROR_CANTOPEN And _
            lReturn <> ERROR_CANTREAD And lReturn <> ERROR_CANTWRITE _
            And lReturn <> ERROR_OUTOFMEMORY And lReturn <> ERROR_INVALID_PARAMETER _
            And lReturn <> ERROR_ACCESS_DENIED And lReturn <> ERROR_NO_MORE_ITEMS
            
            ReDim Preserve Key(EnumKeys)
            
            Key(EnumKeys) = Left$(sName, retVal)
                        
            EnumKeys = EnumKeys + 1
            sName = String(255, 0)
            retVal = Len(sName)
            
            lReturn = RegEnumKeyEx(mainKey, EnumKeys, sName, retVal, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&)
        Wend
    
        RegCloseKey mainKey
    Else
        EnumKeys = -1
    End If
End Function

Public Function HaveSubkey(ByVal sPath As String) As Boolean
Attribute HaveSubkey.VB_Description = "Checks for subkeys in some key..."
    Dim sName As String, retVal As Long, SubKeyCount As Long
    
    hKey = GetKeys(sPath, sKey)
    
    If (RegOpenKey(hKey, sKey, mainKey) = ERROR_SUCCESS) Then 'try to open key

        SubKeyCount = 0
        sName = String(255, 0)
        retVal = Len(sName)
        HaveSubkey = False
        
        lReturn = RegEnumKeyEx(mainKey, SubKeyCount, sName, retVal, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&)
        
        Do While lReturn <> ERROR_BADDB And lReturn <> ERROR_BADKEY And lReturn <> ERROR_CANTOPEN And _
            lReturn <> ERROR_CANTREAD And lReturn <> ERROR_CANTWRITE _
            And lReturn <> ERROR_OUTOFMEMORY And lReturn <> ERROR_INVALID_PARAMETER _
            And lReturn <> ERROR_ACCESS_DENIED And lReturn <> ERROR_NO_MORE_ITEMS
            HaveSubkey = True
            Exit Do
        Loop
    
        RegCloseKey mainKey
    Else
        HaveSubkey = False
    End If
End Function

Public Function CreateValue(ByVal sPath As String, ByVal sName As String, ByVal nType As rcRegType) As Long
    Select Case nType
            Case rcRegType.REG_SZ: CreateValue = WriteString(sPath, sName, vbNullChar)
            Case rcRegType.REG_BINARY: CreateValue = WriteBinary(sPath, sName, "")
            Case rcRegType.REG_DWORD: CreateValue = WriteDWORD(sPath, sName, 0)
            Case rcRegType.REG_MULTI_SZ: CreateValue = WriteString(sPath, sName, vbNullChar, REG_MULTI_SZ)
            Case rcRegType.REG_EXPAND_SZ: CreateValue = WriteString(sPath, sName, vbNullChar, REG_EXPAND_SZ)
            Case Else: CreateValue = 0: Exit Function
    End Select
End Function

Public Function WriteString(ByVal sPath As String, ByVal sName As String, ByVal sValue As String, Optional ByVal sType As rcRegType = rcRegType.REG_SZ) As Long
Attribute WriteString.VB_Description = "Writes data of string type..."
    If (KeyExists(sPath) = False) Then
        If (createNoExists = True) Then
            CreateKey sPath
        Else
            WriteString = 0
            Exit Function
        End If
    End If
    
    hKey = GetKeys(sPath, sKey)

    If (sName = "@") Then sName = ""

    If (RegOpenKeyEx(hKey, sKey, 0, KEY_WRITE, mainKey) = ERROR_SUCCESS) Then
        If (RegSetValueEx(mainKey, sName, 0, sType, ByVal sValue, Len(sValue)) = ERROR_SUCCESS) Then
            RegCloseKey mainKey
            WriteString = mainKey
        Else
            WriteString = 0
      End If
    Else
         WriteString = 0
    End If
End Function

Public Function ReadString(ByVal sPath As String, ByVal sName As String, Optional sDefault As String = vbNullChar) As String
Attribute ReadString.VB_Description = "Reads data of string type..."
    Dim sData As String, lDuz As Long
    
    hKey = GetKeys(sPath, sKey)
    
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
        sData = String(255, 0)
        lDuz = 255
        
        If (RegQueryValueEx(mainKey, sName, 0, REG_SZ, sData, lDuz) = ERROR_SUCCESS) Then
            RegCloseKey mainKey
            sData = Left$(sData, lDuz)
            If Len(sData) > 0 Then ReadString = Left$(sData, Len(sData) - 1) Else: ReadString = sDefault
        Else
            ReadString = sDefault
        End If
    Else
        ReadString = sDefault
    End If
End Function

Public Function WriteDWORD(ByVal sPath As String, ByVal sName As String, ByVal lValue As Long) As Long
Attribute WriteDWORD.VB_Description = "Writes data of dword type..."
    If (KeyExists(sPath) = False) Then
        If (createNoExists = True) Then
            CreateKey sPath
        Else
            WriteDWORD = 0
            Exit Function
        End If
    End If
    hKey = GetKeys(sPath, sKey)
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_WRITE, mainKey) = ERROR_SUCCESS) Then
        If (RegSetValueExA(mainKey, sName, 0, REG_DWORD, lValue, 4) = ERROR_SUCCESS) Then
            RegCloseKey mainKey
            WriteDWORD = mainKey
        Else
            WriteDWORD = 0
        End If
    Else
        WriteDWORD = 0
    End If
End Function

Public Function ReadDWORD(ByVal sPath As String, ByVal sName As String, Optional lDefault As Double = -1) As Long
Attribute ReadDWORD.VB_Description = "Reads data of dword type..."
    Dim lData As Long
    hKey = GetKeys(sPath, sKey)
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
        If (RegQueryValueExA(mainKey, sName, 0, REG_DWORD, lData, 4) = ERROR_SUCCESS) Then
            RegCloseKey mainKey
            ReadDWORD = lData
        Else
            ReadDWORD = lDefault
        End If
    Else
        ReadDWORD = lDefault
    End If
End Function

Public Function WriteBinary(ByVal sPath As String, ByVal sName As String, _
                                                   ByVal sValue As String) As Long
Attribute WriteBinary.VB_Description = "Writes data of binary type..."
    Dim L As Long, lDuz As Long, B() As Byte
    
    If (KeyExists(sPath) = False) Then
        If (createNoExists = True) Then
            CreateKey sPath
        Else
            WriteBinary = 0
            Exit Function
        End If
    End If

    hKey = GetKeys(sPath, sKey)
    
    sValue = StrToBin(sValue)
    
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_WRITE, mainKey) = ERROR_SUCCESS) Then
      
        lDuz = Len(sValue)
        ReDim B(lDuz) As Byte
      
        For L = 1 To lDuz
            B(L) = Asc(Mid$(sValue, L, 1))
        Next
        
        If (lDuz = 0) Then
            ReDim B(1) As Byte
            B(1) = 0
        End If
        
        If (RegSetValueExB(mainKey, sName, 0, REG_BINARY, B(1), lDuz) = ERROR_SUCCESS) Then
            RegCloseKey mainKey
            WriteBinary = mainKey
        Else
            WriteBinary = 0
        End If
    Else
         WriteBinary = 0
    End If

End Function

Public Function ReadBinary(ByVal sPath As String, ByVal sName As String, Optional sDefault As String = vbNullString, Optional ByVal convertTo As convertBINType = BIN_Bin, Optional ByVal posBegin As Long = 1, Optional ByVal totalBytes As Long = 0) As Variant
    Dim arrData() As Byte, lDuz As Long
    Dim sData As String, strData As String, cData As Integer, nI As Long
    ReDim arrData(0 To 1) As Byte
    arrData(0) = 0
    hKey = GetKeys(sPath, sKey)
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
        lDuz = 1
        RegQueryValueEx mainKey, sName, 0, REG_BINARY, 0, lDuz
        sData = String(lDuz, 0)
        If lDuz > 0 Then
            RegQueryValueEx mainKey, sName, 0, REG_BINARY, sData, lDuz
            RegCloseKey mainKey
            If totalBytes = 0 Then totalBytes = lDuz Else sData = Mid(sData, posBegin, totalBytes)
            Select Case convertTo
                Case BIN_Array
                    ReDim arrData(totalBytes - 1) As Byte
                    For nI = 1 To totalBytes
                        arrData(nI - 1) = Asc(Mid$(sData, nI, 1))
                    Next
                    ReadBinary = arrData
                Case BIN_String
                    ReadBinary = sData
                Case BIN_Dot
                    strData = ""
                    For nI = 1 To totalBytes
                        cData = Asc(Mid$(sData, nI, 1))
                        If cData < 33 Or (cData > 126 And cData < 144) Or (cData > 147 And cData < 161) Then
                            strData = strData & "."
                        Else
                            strData = strData & Mid$(sData, nI, 1)
                        End If
                        If nI > 0 And nI Mod 8 = 0 Then strData = strData & Chr$(13)
                    Next
                    ReadBinary = strData
                Case BIN_Bin
                    ReadBinary = Trim$(BinToStr(sData))
            End Select
        Else
            ReadBinary = sDefault
        End If
    Else
        ReadBinary = sDefault
    End If
End Function

Public Function KillValue(ByVal sPath As String, ByVal sName As String) As Long
Attribute KillValue.VB_Description = "Deletes value from registry..."

    hKey = GetKeys(sPath, sKey)
    
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_ALL_ACCESS, mainKey) = ERROR_SUCCESS) Then
        RegDeleteValue mainKey, sName
        RegCloseKey mainKey
        KillValue = mainKey
    Else
        KillValue = 0
    End If
    
End Function

Public Function ValueExists(ByVal sPath As String, ByVal sName As String) As Boolean
Attribute ValueExists.VB_Description = "Checks if some value exists in registry"
    
    hKey = GetKeys(sPath, sKey)
    
    Dim sData As String
    
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
        
        If (RegQueryValueEx(mainKey, sName, 0, 0, sData, 1) = ERROR_SUCCESS) Then
            RegCloseKey mainKey
            ValueExists = True
        Else
            ValueExists = False
        End If
    Else
        ValueExists = False
    End If
    
End Function

Public Function EnumValues(ByVal sPath As String, ByRef sValue() As String, ByRef sData() As Variant, Optional ByVal OnlyType As rcRegType = -1) As Long
Attribute EnumValues.VB_Description = "Enumerates all values from specified key..."
    Const lengthVars = 2000
    Dim rName As String
    Dim rData As Long
    Dim mKey As Long, rType As Long, Cnt As Long
    Dim RetData As Long, retValue As Long
    
    hKey = GetKeys(sPath, sKey)
    If RegOpenKey(hKey, sKey, mKey) = ERROR_SUCCESS Then
        Cnt = 0
        Erase sValue
        Erase sData
        Do
            rName = String(lengthVars, 0)
            retValue = lengthVars
            RetData = lengthVars
            lReturn = RegEnumValue(mKey, Cnt, ByVal rName, retValue, 0, rType, ByVal rData, RetData)
            Cnt = Cnt + 1
            If lReturn = ERROR_SUCCESS Then
                If (OnlyType = -1) Or (OnlyType = rType) Then
                    ReDim Preserve sValue(EnumValues) As String
                    ReDim Preserve sData(EnumValues) As Variant
                    rName = Trim$(Left$(rName, retValue))
                    sValue(EnumValues) = rName
                    Select Case rType
                        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
                            sData(EnumValues) = ReadString(sPath, rName)
                        Case REG_BINARY
                            sData(EnumValues) = ReadBinary(sPath, rName)
                        Case REG_DWORD
                            sData(EnumValues) = ReadDWORD(sPath, rName)
                    End Select
                    EnumValues = EnumValues + 1
                End If
            End If
        Loop While lReturn = ERROR_SUCCESS
        RegCloseKey hKey
    Else
        EnumValues = -1
    End If
End Function

Public Sub ExportToReg(ByVal sRegFile As String, ByVal sPath As String, Optional IncludeSubKeys As Boolean = True)
    Dim fFile As Long
    
    sPath = Mid$(sPath, Len("My Computer\") + 1)
    If sPath = "" Then Exit Sub
    
    fFile = FreeFile
    Open sRegFile For Output As #fFile
        Print #fFile, "Windows Registry Editor Version 5.00"
        Print #fFile, vbCrLf & "[" & sPath & "]"
        GenKey fFile, sPath, IncludeSubKeys
    Close #fFile
End Sub

Sub GenKey(ByVal fFile As Long, ByVal sPath As String, Optional IncludeSubKeys As Boolean = True)
    Dim nK As Long, nV As Long, nT As Byte
    Dim nKeys As Long, nValues As Long, sK As String
    Dim Keys() As String, Values() As String, Datas() As Variant
    Dim sVal As String, nS As Long
    Dim arrTypes()
    
    arrTypes = Array(rcRegType.REG_SZ, rcRegType.REG_BINARY, rcRegType.REG_DWORD, rcRegType.REG_MULTI_SZ, rcRegType.REG_EXPAND_SZ)
    
    nKeys = EnumKeys(sPath, Keys())
    For nT = 0 To 4
        nValues = EnumValues(sPath, Values, Datas, arrTypes(nT))
        For nV = 0 To nValues - 1
            sVal = ""
            Select Case arrTypes(nT)
                Case rcRegType.REG_SZ
                    Datas(nV) = Chr$(34) & Datas(nV) & Chr$(34)
                Case rcRegType.REG_BINARY
                    Datas(nV) = "hex:" & Replace$(Datas(nV), " ", ",", , , vbTextCompare)
                Case rcRegType.REG_DWORD
                    Datas(nV) = "dword:" & frmRegEdit.Dbl2Hex(Datas(nV), 8)
                Case rcRegType.REG_MULTI_SZ, rcRegType.REG_EXPAND_SZ
                    sVal = Datas(nV)
                    Datas(nV) = ""
                    For nS = 1 To Len(sVal)
                        Datas(nV) = Datas(nV) & IIf(nS > 1, ",", "") & frmRegEdit.Dbl2Hex(Asc(Mid(sVal, nS, 1)), 2) & ",00"
                    Next
                    Datas(nV) = "hex(" & IIf(arrTypes(nT) = rcRegType.REG_MULTI_SZ, 7, 2) & "):" & Datas(nV) & IIf(Datas(nV) <> "", ",", "") & "00,00"
            End Select
            Print #fFile, Chr$(34) & Values(nV) & Chr$(34) & "=" & Datas(nV)
        Next nV
    Next nT
    If nKeys > 0 Then
        For nK = 0 To nKeys - 1
            sK = sPath & "\" & CStr(Keys(nK))
            Print #fFile, vbCrLf & "[" & sK & "]" & vbCrLf
            If IncludeSubKeys Then GenKey fFile, sK, IncludeSubKeys
        Next
    End If
End Sub

Public Sub ImportFromReg(ByVal sRegFile As String)
    Dim fFile As Long
    Dim Lines() As String, nL As Long, nP As Long
    Dim cLine As String, curKey As String, vName As String, vData As String
    
    CreateKeyIfDoesntExists = True
    
    fFile = FreeFile
    Open sRegFile For Binary As #fFile
        Lines = Split(Replace$(Replace$(Replace$(Input(LOF(fFile), #fFile), Chr$(0), ""), Chr$(&HFF) & Chr$(&HFE), ""), "\" & vbCrLf, ""), vbCrLf)
    Close #fFile
    
    If (UCase$(Lines(0)) <> UCase("Windows Registry Editor Version 5.00")) Then
        MsgBox "Cannot import " & sRegFile & ": The specified file is not a registry file. You can import only registry files.", vbOKOnly + vbCritical, frmRegEdit.Caption
        Exit Sub
    End If

    curKey = ""
    cLine = ""
    For nL = 1 To UBound(Lines)
        If Left$(Lines(nL), 1) = "[" Then
            curKey = Mid$(Lines(nL), 2, Len(Lines(nL)) - 2)
            CreateKey curKey
        Else
            cLine = Replace$(Trim$(Lines(nL)), ChrW$(34), "")
            If curKey <> "" And cLine <> "" Then
                nP = InStr(1, cLine, "=", vbTextCompare)
                vName = Left$(cLine, nP - 1)
                vData = Mid$(cLine, nP + 1)
                If Len(vData) > 4 Then
                    If Left$(vData, 4) = "hex:" Then
                        vData = Replace$(Mid$(vData, 5), ",", " ", , , vbTextCompare)
                        WriteBinary curKey, vName, vData
                    Else
                        If Left$(vData, 6) = "dword:" Then
                            vData = Mid$(vData, 7)
                            WriteDWORD curKey, vName, CDbl("&H" & vData)
                        Else
                            If Left$(vData, 7) = "hex(7):" Or Left$(vData, 7) = "hex(2):" Then
                                vData = Replace$(vData, ",", " ", , , vbTextCompare)
                                WriteString curKey, vName, Replace$(StrToBin(Mid(vData, 8)), Chr$(0), ""), IIf(Left$(vData, 7) = "hex(7):", REG_MULTI_SZ, REG_EXPAND_SZ)
                            Else
                                WriteString curKey, vName, vData, REG_SZ
                            End If
                        End If
                    End If
                Else
                    WriteString curKey, vName, vData, REG_SZ
                End If
            End If
        End If
    Next nL
End Sub

Public Function StrToBin(sBin As String) As String
    Dim sRes As String, nPos As Long
    sBin = Trim$(Replace(sBin, " ", vbNullString))
    sRes = ""
    For nPos = 1 To Len(sBin) Step 2
        sRes = sRes & ChrW$("&H" & Mid(sBin, nPos, 2))
    Next nPos
    StrToBin = sRes
End Function

Public Function BinToStr(ByVal sStr As String, Optional ByVal isTruncate As Boolean = True) As String
Attribute BinToStr.VB_Description = "Converts binary data into string..."
    Const maxLen = 128
    Dim retVal As String
    Dim nI As Integer, sByte As String
    If Len(sStr) = 0 Then GoTo zeroLength
    If isTruncate Then If Len(sStr) > maxLen Then sStr = Left$(sStr, maxLen)
    retVal = vbNullString
    For nI = 1 To Len(sStr)
        sByte = Mid$(sStr, nI, 1)
        retVal = retVal & " "
        If sByte = vbNullChar Then
            retVal = retVal & "00"
        Else
            If Asc(sByte) < &HF Then retVal = retVal & "0"
            retVal = retVal & Hex$(Asc(sByte))
        End If
    Next nI
zeroLength:
    BinToStr = retVal
End Function

Public Function isBinValid(ByVal sBin As String) As Boolean
Attribute isBinValid.VB_Description = "Checks if given binary data is in valid hex format (used for writting binary)"
    Dim Z As Long
    sBin = Trim$(UCase$(Replace(sBin, " ", vbNullString)))
    If Len(sBin) = 0 Then GoTo zeroLength
    For Z = 1 To Len(sBin)
        If InStr(1, Mid$(sBin, Z, 1), "0123456789ABCDEF ", 1) < 1 Then
zeroLength:
           isBinValid = False
           Exit Function
        End If
    Next
    isBinValid = True
End Function

Function renameKey(ByVal sKeySource As String, ByVal sNewName As String) As Boolean
    Dim hKeySource As Long
    Dim hKeyDestination As Long
    Dim nResult As Long
    Dim sFile As String
On Error GoTo ErrHandler
    sNewName = Mid(sKeySource, 1, InStrRev(sKeySource, "\")) & sNewName
    hKey = GetKeys(sNewName, sKey)
    sNewName = sKey
    hKey = GetKeys(sKeySource, sKey)

    setBackupAndRestorePriviliges
    sFile = App.Path & "\" & App.EXEName & "."
    If Len(Dir(sFile)) > 0 Then Kill sFile
    
    'Check for exist of source
    nResult = RegOpenKeyEx(hKey, sKey, 0&, KEY_ALL_ACCESS, hKeySource)
    If nResult <> ERROR_SUCCESS Then GoTo ErrHandler
    'Save old informations
    nResult = RegSaveKey(hKeySource, sFile, 0&)
    If nResult <> ERROR_SUCCESS Then GoTo ErrHandler
    'Check for exist of destination
    nResult = RegOpenKeyEx(hKey, sNewName, 0&, KEY_ALL_ACCESS, hKeyDestination)
    If nResult <> ERROR_FILE_NOT_FOUND Then GoTo ErrHandler
    'Try to delete source & re-check for self
    If KillKey(sKeySource) = 0 Or KeyExists(sKeySource) Then GoTo ErrHandler
    'If delete successfull then create new key
    nResult = RegCreateKeyEx(hKey, sNewName, 0&, vbNullString, 0&, KEY_ALL_ACCESS, 0&, hKeyDestination, 0&)
    If nResult <> ERROR_SUCCESS Then GoTo ErrHandler
    'and restore old informations to new key
    nResult = RegRestoreKey(hKeyDestination, sFile, REG_FORCE_RESTORE)
    
    resetBackupAndRestorePriviliges
    If Len(Dir(sFile)) > 0 Then Kill sFile
    renameKey = True
    Exit Function

ErrHandler:
    If Len(Dir(sFile)) > 0 Then Kill sFile
    RegCloseKey hKeySource
    RegCloseKey hKeyDestination
    renameKey = False
End Function

Private Sub setBackupAndRestorePriviliges()
    Dim nResult As Long
On Error GoTo ErrHandler
    
    nResult = OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVLEGES Or TOKEN_QUERY, m_hToken)
    nResult = LookupPrivilegeValue(vbNullString, SE_RESTORE_NAME, m_RestoreLuid)
    nResult = LookupPrivilegeValue(vbNullString, SE_BACKUP_NAME, m_BackupLuid)
    
    m_TP.PrivilegeCount = 2
    m_TP.Privileges(0).pLuid = m_RestoreLuid
    m_TP.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    m_TP.Privileges(1).pLuid = m_BackupLuid
    m_TP.Privileges(1).Attributes = SE_PRIVILEGE_ENABLED
    
    nResult = AdjustTokenPrivileges(m_hToken, vbFalse, m_TP, Len(m_TP), 0&, 0&)
ErrHandler:
End Sub

Private Sub resetBackupAndRestorePriviliges()
    Dim nResult As Long
On Error GoTo ErrHandler
    nResult = AdjustTokenPrivileges(m_hToken, vbTrue, m_TP, Len(m_TP), 0&, 0&)
ErrHandler:
End Sub

Private Sub Class_Initialize()
    CreateKeyIfDoesntExists = True
End Sub
