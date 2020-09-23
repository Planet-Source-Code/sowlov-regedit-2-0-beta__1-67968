VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRegEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "Registry Editor"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegEdit.frx":0000
   ScaleHeight     =   5205
   ScaleWidth      =   6825
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdDialog 
      Left            =   1065
      Top             =   1635
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Registration Files (*.reg)|*.reg"
      FilterIndex     =   2
      MaxFileSize     =   1024
   End
   Begin VB.PictureBox pbResize 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   3645
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4860
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   0
      Width           =   15
   End
   Begin MSComctlLib.ImageList ilImages 
      Left            =   405
      Top             =   1650
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEdit.frx":014A
            Key             =   "MyComputer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEdit.frx":049C
            Key             =   "Opened"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEdit.frx":078E
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEdit.frx":0AB0
            Key             =   "val1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEdit.frx":0E02
            Key             =   "val2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEdit.frx":1154
            Key             =   "Finding"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegEdit.frx":1DA6
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStastus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   4905
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView tvKeys 
      Height          =   4905
      Left            =   0
      TabIndex        =   1
      Top             =   -15
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   8652
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   132
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ilImages"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvValues 
      Height          =   4905
      Left            =   3675
      TabIndex        =   0
      Top             =   -15
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   8652
      View            =   3
      Arrange         =   2
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ilImages"
      SmallIcons      =   "ilImages"
      ColHdrIcons     =   "ilImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuImport 
         Caption         =   "Import..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLHive 
         Caption         =   "Load Hive..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuULHive 
         Caption         =   "Unload Hive..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect Network Registry..."
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect Network Registry..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuModify 
         Caption         =   "Modify"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuModifyBIN 
         Caption         =   "Modify Binary Data"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExpand 
         Caption         =   "Expand"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCollapse 
         Caption         =   "Collapse"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep00 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Begin VB.Menu mnuNewKey 
            Caption         =   "Key"
         End
         Begin VB.Menu mnuSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNewString 
            Caption         =   "String value"
         End
         Begin VB.Menu mnuNewBinary 
            Caption         =   "Binary value"
         End
         Begin VB.Menu mnuNewDWORD 
            Caption         =   "DWORD value"
         End
         Begin VB.Menu mnuNewMultiString 
            Caption         =   "Multi-String value"
         End
         Begin VB.Menu mnuNewExpString 
            Caption         =   "Explandabled String value"
         End
      End
      Begin VB.Menu mnuFindTree 
         Caption         =   "Find..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExport 
         Caption         =   "Export"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPermission 
         Caption         =   "Permission..."
      End
      Begin VB.Menu mnuSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy Key Name"
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDisconnect 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuStatusBar 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep08 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSplit 
         Caption         =   "Split"
      End
      Begin VB.Menu mnuSep09 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDspBinData 
         Caption         =   "Display Binary Data"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "Favorites"
      Begin VB.Menu mnuAddFavorites 
         Caption         =   "Add to Favorites"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRemoveFavorite 
         Caption         =   "Remove Favorite"
      End
      Begin VB.Menu mnuSep07 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFavIdx 
         Caption         =   "Favorite Index..."
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help Topics"
         HelpContextID   =   1
      End
      Begin VB.Menu mnuSep06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Registry Editor"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "Tools"
      Begin VB.Menu mnuTools 
         Caption         =   "Tools"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "Options..."
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "frmRegEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Const sDefaultName = "(Default)"
Const sDefaultValue = "(value not set)"
Const sDefaultZERO = "(zero-length binary value)"
Const sDefaultWORD = "0x00000000 (0)"

Public clsReg As clsRegistryAccess

Dim isDrag As Boolean
Dim isRMouse As Boolean
Dim isDblClk As Boolean
Dim oldName As MSComctlLib.Node
Dim mX As Integer, mY As Integer
Dim curPos As String
Public regPath As String

Public envFindFlag As Double
Dim envLastKey As String
Dim envView As String
Dim envColumnWidth(2) As Long

Public fPath As String
Public fValue As String

Const strHEX = "0123456789ABCDEF"
Public Function Dbl2Hex(ByVal nDbl As Double, Optional ByVal nZero As Integer = 0) As String
    Dim nMod As Byte
    Dbl2Hex = ""
    While nDbl > 0
        'Proccess for the great number
        nMod = nDbl - (Int(nDbl / 16) * 16)
        Dbl2Hex = Mid(strHEX, nMod + 1, 1) & Dbl2Hex
        nDbl = Int(nDbl / 16)
    Wend
    If nZero > 0 Then
        For nMod = Len(Dbl2Hex) + 1 To nZero
            Dbl2Hex = "0" & Dbl2Hex
        Next
    End If
End Function
Public Function Hex2Dbl(sHex As String) As Double
    Dim nPos As Integer
    Hex2Dbl = 0
    For nPos = Len(sHex) To 1 Step -1
        Hex2Dbl = Hex2Dbl + (CByte("&H" & Mid(strHEX, InStr(1, strHEX, Mid(sHex, nPos, 1)), 1)) * (16 ^ (Len(sHex) - nPos)))
    Next nPos
End Function

Private Sub openPopup(ByVal x As Integer, ByVal y As Integer)
On Error GoTo exitSub
    hideAllSubEditMenu
    If curPos = "tv" Then
        With tvKeys.SelectedItem
            If .Key = "" Then
                mnuSep02.Visible = True
                setNewsSub
                mnuPermission.Visible = True
                mnuDelete.Visible = True
                mnuRename.Visible = True
                mnuDelete.Enabled = InStr(1, Replace(.FullPath, "My Computer\", ""), "\") > 0
                mnuRename.Enabled = mnuDelete.Enabled
                mnuSep04.Visible = True
                mnuFind.Visible = True
                mnuSep05.Visible = True
                mnuCopy.Visible = True
            Else
                mnuSep02.Visible = True
                mnuEditExport.Visible = True
                mnuSep15.Visible = True
                mnuEditDisconnect.Visible = True
                mnuNew.Visible = False
            End If
            mnuSep03.Visible = mnuDelete.Visible
            mnuExpand.Visible = Not .Expanded
            mnuExpand.Enabled = .Children > 0
            mnuCollapse.Visible = .Expanded
            mnuCollapse.Enabled = .Children > 0
            PopupMenu mnuEdit, , , , IIf(mnuExpand.Visible, mnuExpand, mnuCollapse)
        End With
    Else
        setNewsSub
        PopupMenu mnuEdit
    End If
exitSub:
    sbStastus.SimpleText = tvKeys.SelectedItem.FullPath
End Sub

Public Function explainPath(ByVal oPath As String) As String
    oPath = Replace(oPath, "My Computer\", "")
    oPath = Replace(oPath, "HKEY_CLASSES_ROOT", "HKCR")
    oPath = Replace(oPath, "HKEY_CURRENT_USER", "HKCU")
    oPath = Replace(oPath, "HKEY_LOCAL_MACHINE", "HKLM")
    oPath = Replace(oPath, "HKEY_USERS", "HKUS")
    oPath = Replace(oPath, "HKEY_CURRENT_CONFIG", "HKCC")
    oPath = Replace(oPath, "HKEY_PERFORMANCE_DATA", "HKPD")
    oPath = Replace(oPath, "HKEY_DYN_DATA", "HKDD")
    explainPath = oPath
End Function

Private Sub addnewSubKeys(ByVal Node As MSComctlLib.Node)
    Dim sKeys() As String, nKeys As Long
    Dim nCnt As Long
    If Node.Children = 0 And clsReg.HaveSubkey(explainPath(Node.FullPath)) Then
        nKeys = clsReg.EnumKeys(explainPath(Node.FullPath), sKeys)
        If nKeys > 0 Then
            For nCnt = 0 To nKeys - 1
                tvKeys.Nodes.Add Node, tvwChild, , sKeys(nCnt), "Closed", "Opened"
            Next nCnt
        End If
    End If
End Sub

Public Sub addnewValues(ByVal Node As MSComctlLib.Node)
    Dim sNames() As String, sValues() As Variant, nValues As Long, nVal As Long
    Dim arrTypes(), nArr As Integer, hadDefault As Boolean
    lvValues.ListItems.Clear
    If Node.Key <> "" Then Exit Sub
    hadDefault = False
    arrTypes = Array(rcRegType.REG_SZ, rcRegType.REG_BINARY, rcRegType.REG_DWORD, rcRegType.REG_MULTI_SZ, rcRegType.REG_EXPAND_SZ)
    For nArr = 0 To 4
        nValues = clsReg.EnumValues(explainPath(Node.FullPath), sNames, sValues, arrTypes(nArr))
        For nVal = 0 To nValues - 1
            With lvValues.ListItems.Add(, , Trim(sNames(nVal)))
                .Tag = .Text
                Select Case arrTypes(nArr)
                    Case rcRegType.REG_BINARY, rcRegType.REG_DWORD
                        .Icon = "val2": .SmallIcon = "val2"
                    Case Else
                        .Icon = "val1": .SmallIcon = "val1"
                End Select
                With .ListSubItems.Add(, , "")
                    Select Case arrTypes(nArr)
                        Case rcRegType.REG_SZ: .Text = "REG_SZ"
                        Case rcRegType.REG_BINARY: .Text = "REG_BINARY"
                        Case rcRegType.REG_DWORD: .Text = "REG_DWORD"
                        Case rcRegType.REG_MULTI_SZ: .Text = "REG_MULTI_SZ"
                        Case rcRegType.REG_EXPAND_SZ: .Text = "REG_EXPAND_SZ"
                    End Select
                End With
                With .ListSubItems.Add(, , sValues(nVal))
                    Select Case arrTypes(nArr)
                            Case rcRegType.REG_SZ:
                                .Text = sValues(nVal)
                            Case rcRegType.REG_BINARY:
                                If sValues(nVal) <> "" Then
                                    .Text = UCase(sValues(nVal))
                                    If Len(.Text) > 128 Then .Text = Left(.Text, 128) & "..."
                                Else
                                    .Text = sDefaultZERO
                                End If
                            Case rcRegType.REG_DWORD:
                                .Text = Dbl2Hex(sValues(nVal))
                                While Len(.Text) < 8
                                    .Text = "0" & .Text
                                Wend
                                .Text = "0x" & .Text & " (" & sValues(nVal) & ")"
                            Case rcRegType.REG_MULTI_SZ, rcRegType.REG_EXPAND_SZ:
                                .Text = CStr(sValues(nVal))
                    End Select
                End With
                If .Text = "" Then
                    .Text = sDefaultName
                    If .ListSubItems(1).Text = "" Then .ListSubItems(1).Text = "REG_SZ"
                    If .ListSubItems(2).Text = "" Then .ListSubItems(2).Text = sDefaultValue
                    hadDefault = True
                End If
            End With
        Next nVal
    Next nArr
    If Not hadDefault Then
        With lvValues.ListItems.Add(, sDefaultName, sDefaultName, "val1", "val1")
            .ListSubItems.Add , , "REG_SZ"
            .ListSubItems.Add , , sDefaultValue
        End With
    End If
End Sub

Private Sub addnewValue(ByVal vType As rcRegType)
On Error Resume Next
    Dim nIdx As Integer
    nIdx = 1
    While clsReg.ValueExists(explainPath(sbStastus.SimpleText), "New Value #" & nIdx)
        nIdx = nIdx + 1
    Wend
    clsReg.CreateValue explainPath(sbStastus.SimpleText), "New Value #" & nIdx, vType
    With lvValues.ListItems.Add(, , "New Value #" & nIdx)
        Select Case vType
            Case rcRegType.REG_BINARY, rcRegType.REG_DWORD
                .Icon = "val2": .SmallIcon = "val2"
            Case Else
                .Icon = "val1": .SmallIcon = "val1"
        End Select
        With .ListSubItems.Add(, , "")
            Select Case vType
                Case rcRegType.REG_SZ: .Text = "REG_SZ"
                Case rcRegType.REG_BINARY: .Text = "REG_BINARY"
                Case rcRegType.REG_DWORD: .Text = "REG_DWORD"
                Case rcRegType.REG_MULTI_SZ: .Text = "REG_MULTI_SZ"
                Case rcRegType.REG_EXPAND_SZ: .Text = "REG_EXPAND_SZ"
            End Select
        End With
        With .ListSubItems.Add()
            Select Case vType
                Case rcRegType.REG_BINARY: .Text = sDefaultZERO
                Case rcRegType.REG_DWORD: .Text = sDefaultWORD
            End Select
        End With
        lvValues.SelectedItem.Selected = False
        .Selected = True
        lvValues.StartLabelEdit
    End With
End Sub

Private Sub setNewsSub()
    Dim vSet As Boolean
    vSet = tvKeys.SelectedItem.Key = ""
    mnuNewKey.Enabled = vSet
    mnuNewString.Enabled = vSet
    mnuNewBinary.Enabled = vSet
    mnuNewDWORD.Enabled = vSet
    mnuNewMultiString.Enabled = vSet
    mnuNewExpString.Enabled = vSet
End Sub

Private Sub hideAllSubEditMenu()
    mnuModify.Visible = False
    mnuModifyBIN.Visible = False
    mnuCollapse.Visible = False
    mnuExpand.Visible = False
    mnuSep00.Visible = False
    mnuNew.Visible = True
    setNewsSub
    mnuSep02.Visible = False
    mnuEditExport.Visible = False
    mnuPermission.Visible = False
    mnuSep03.Visible = False
    mnuDelete.Visible = False
    mnuRename.Visible = False
    mnuSep04.Visible = False
    mnuFind.Visible = False
    mnuFindNext.Visible = False
    mnuSep05.Visible = False
    mnuCopy.Visible = False
    mnuSep15.Visible = False
    mnuEditDisconnect.Visible = False
End Sub

Private Sub proccessMenu()
    If isRMouse Then Exit Sub
On Error Resume Next
    mnuModify.Visible = curPos = "lv"
    mnuModifyBIN.Visible = curPos = "lv"
    mnuSep00.Visible = curPos = "lv"
    mnuNew.Visible = True
    setNewsSub
    mnuSep02.Visible = True
    mnuEditExport.Visible = False
    mnuPermission.Visible = True
    mnuSep03.Visible = True
    mnuDelete.Visible = True
    mnuDelete.Enabled = InStr(1, Replace(tvKeys.SelectedItem.FullPath, "My Computer\", ""), "\") > 0
    mnuRename.Visible = True
    mnuRename.Enabled = mnuDelete.Enabled
    mnuSep04.Visible = True
    mnuFind.Visible = True
    mnuFindNext.Visible = True
    mnuSep05.Visible = True
    mnuCopy.Visible = True
    mnuSep15.Visible = False
    mnuEditDisconnect.Visible = False
    mnuAddFavorites.Enabled = tvKeys.SelectedItem.Index > 1
    getFavorites
    sbStastus.SimpleText = tvKeys.SelectedItem.FullPath
End Sub

Private Function isMulti() As Boolean
    Dim nItems As Integer, nIndex As Integer
    nItems = 0
    For nIndex = 1 To lvValues.ListItems.Count
        If lvValues.ListItems(nIndex).Selected Then nItems = nItems + 1
    Next nIndex
    isMulti = nItems > 1
End Function

Private Sub proccessKeyCode(ByVal KeyCode As Integer)
On Error GoTo exitSub
    Select Case KeyCode
        Case 113: If mnuRename.Enabled Then mnuRename_Click
    End Select
exitSub:
End Sub

Public Sub getFavorites()
    Dim sNames() As String, sValues() As Variant, nValues As Long, nVal As Long
    For nVal = mnuFavIdx.Count - 1 To 1 Step -1
        Unload mnuFavIdx(nVal)
    Next nVal
    nValues = clsReg.EnumValues(regPath & "\Favorites", sNames, sValues, REG_SZ)
    For nVal = 0 To nValues - 1
        Load mnuFavIdx(mnuFavIdx.Count)
        With mnuFavIdx(mnuFavIdx.Count - 1)
            .Caption = Trim(sNames(nVal))
            .Tag = sValues(nVal)
            .Visible = True
        End With
    Next nVal
    mnuSep07.Visible = mnuFavIdx.Count > 1
    mnuAddFavorites.Enabled = InStr(1, tvKeys.SelectedItem.FullPath, "\") > 0
    mnuRemoveFavorite.Enabled = mnuSep07.Visible
End Sub

Private Function findSubNode(parNode As MSComctlLib.Node, ByVal chdName As String) As Long
    Dim nIdx As Integer
    Dim nI As Integer, nChd As MSComctlLib.Node
    nIdx = 0
    If parNode.Children > 0 Then
        Set nChd = parNode.Child
        If UCase(nChd.Text) = UCase(Left(chdName, Len(nChd.Text))) Then
            nIdx = nChd.Index
            Set parNode = nChd
        Else
            For nI = 1 To parNode.Children - 1
                Set nChd = nChd.Next
                If UCase(nChd.Text) = UCase(Left(chdName, Len(nChd.Text))) Then
                    nIdx = nChd.Index
                    Set parNode = nChd
                    Exit For
                End If
            Next nI
        End If
    End If
    findSubNode = nIdx
End Function

Public Sub gotoAKey(ByVal sName As String)
On Error Resume Next
    Dim curNode As MSComctlLib.Node
    sName = sName & "\"
    sName = Mid(sName, InStr(1, sName, "\", vbTextCompare) + 1)
    Set curNode = tvKeys.Nodes(1)
    Do While InStr(1, sName, "\", vbTextCompare) > 0
        If findSubNode(curNode, Left(sName, InStr(1, sName, "\", vbTextCompare) - 1)) > 0 Then
            curNode.Expanded = True
            sName = Mid(sName, InStr(1, sName, "\", vbTextCompare) + 1)
        Else
            Exit Do
        End If
    Loop
    curNode.Selected = True
    addnewValues curNode
    sbStastus.SimpleText = curNode.FullPath
    Set curNode = Nothing
    getFavorites
End Sub

Public Function Hex2Word(ByVal sHex As String) As Double
    Hex2Word = CDbl("&H" & Mid(sHex, 7, 2) & Mid(sHex, 5, 2) & Mid(sHex, 3, 2) & Left(sHex, 2))
End Function

Public Function Word2Hex(ByVal nWord As Double) As String
    Dim sRes As String, nI As Integer
    sRes = Hex(nWord)
    For nI = Len(sRes) + 1 To 8
        sRes = "0" & sRes
    Next nI
    Word2Hex = Right(sRes, 2) & Mid(sRes, 5, 2) & Mid(sRes, 3, 2) & Left(sRes, 2)
End Function

Private Sub Form_Activate()
    Form_Resize
End Sub

Private Sub Form_Load()
    Dim nodRoot As MSComctlLib.Node
    Set clsReg = New clsRegistryAccess
    Const maxArr = 5
    Dim arrTools(maxArr, 1) As String, nI
    arrTools(1, 0) = "HK_CurrentUser -> Run"
    arrTools(1, 1) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
    arrTools(2, 0) = "HK_LocalMachine -> Run"
    arrTools(2, 1) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    arrTools(3, 0) = "UnInstall Program(s)..."
    arrTools(3, 1) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    arrTools(4, 0) = "Windows Policies..."
    arrTools(4, 1) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies"
    arrTools(5, 0) = "Shell Folders"
    arrTools(5, 1) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    For nI = 1 To maxArr
        Load mnuTools(nI)
        With mnuTools(nI)
            .Caption = arrTools(nI, 0)
            .Tag = arrTools(nI, 1)
            .Visible = True
        End With
    Next nI
    
    clsReg.CreateKeyIfDoesntExists = True
    With tvKeys
        Set nodRoot = .Nodes.Add(, tvwRootLines, "\", "My Computer", "MyComputer", "MyComputer")
        .Nodes.Add nodRoot, tvwChild, , "HKEY_CLASSES_ROOT", "Closed", "Opened"
        .Nodes.Add nodRoot, tvwChild, , "HKEY_CURRENT_USER", "Closed", "Opened"
        .Nodes.Add nodRoot, tvwChild, , "HKEY_LOCAL_MACHINE", "Closed", "Opened"
        .Nodes.Add nodRoot, tvwChild, , "HKEY_USERS", "Closed", "Opened"
        .Nodes.Add nodRoot, tvwChild, , "HKEY_CURRENT_CONFIG", "Closed", "Opened"
        .Nodes.Add nodRoot, tvwChild, , "HKEY_PERFORMANCE_DATA", "Closed", "Opened"
        .Nodes.Add nodRoot, tvwChild, , "HKEY_DYN_DATA", "Closed", "Opened"
        nodRoot.Selected = True
        nodRoot.Expanded = True
    End With
    Set nodRoot = Nothing
    
    'Load Configures
    regPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit"
    envFindFlag = clsReg.ReadDWORD(regPath, "FindFlags")
    envLastKey = clsReg.ReadString(regPath, "LastKey")
    envView = clsReg.ReadBinary(regPath, "View")
    clsReg.StrToBin envView
    'windowstate
    Me.WindowState = Hex2Word(Mid(envView, 9, 8))
    '(left,top):(width:height) of mainform
    Me.Left = Me.ScaleX(Hex2Word(Mid(envView, 57, 8)), vbPixels, vbTwips)
    Me.Top = Me.ScaleY(Hex2Word(Mid(envView, 57 + 8, 8)), vbPixels, vbTwips)
    Me.Width = Me.ScaleX(Hex2Word(Mid(envView, 57 + 8 + 8, 8)) - _
                        Hex2Word(Mid(envView, 57, 8)), vbPixels, vbTwips)
    Me.Height = Me.ScaleY(Hex2Word(Mid(envView, 57 + 8 + 8 + 8, 8)) - _
                        Hex2Word(Mid(envView, 57 + 8, 8)), vbPixels, vbTwips)
    'postition of split bar
    pbResize.Left = Me.ScaleX(Hex2Word(Mid(envView, 89, 8)), vbPixels, vbTwips)
    'width of columns
    envColumnWidth(0) = Me.ScaleX(Hex2Word(Mid(envView, 97, 8)), vbPixels, vbTwips)
    envColumnWidth(1) = Me.ScaleX(Hex2Word(Mid(envView, 97 + 8, 8)), vbPixels, vbTwips)
    envColumnWidth(2) = Me.ScaleX(Hex2Word(Mid(envView, 97 + 8 + 8, 8)), vbPixels, vbTwips)
    'statusbar visible
    mnuStatusBar.Checked = Hex2Word(Mid(envView, 121, 8)) = 1
    sbStastus.Visible = mnuStatusBar.Checked

    gotoAKey envLastKey
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If pbResize.Left > Me.ScaleWidth Then pbResize.Left = Me.ScaleX(Me.ScaleX(Me.ScaleWidth, vbTwips, vbPixels) - 100, vbPixels, vbTwips)
    If Me.ScaleX(Me.ScaleWidth, vbTwips, vbPixels) < 300 Then Me.Width = Me.ScaleX(300, vbPixels, vbTwips)
    If Me.ScaleY(Me.ScaleHeight, vbTwips, vbPixels) < 200 Then Me.Height = Me.ScaleY(200, vbPixels, vbTwips)
    tvKeys.Move Me.ScaleLeft - 10, Me.ScaleTop - 10, pbResize.Left - 10, Me.ScaleHeight - IIf(sbStastus.Visible, sbStastus.Height, 0) + 10
    pbResize.Move pbResize.Left, tvKeys.Top, pbResize.Width, tvKeys.Height
    With lvValues
        .Move pbResize.Left + pbResize.Width, tvKeys.Top, Me.ScaleWidth - pbResize.Left - pbResize.Width, tvKeys.Height
        .ColumnHeaders(1).Width = envColumnWidth(0)
        .ColumnHeaders(2).Width = envColumnWidth(1)
        .ColumnHeaders(3).Width = envColumnWidth(2)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save Configures
    clsReg.WriteDWORD regPath, "FindFlags", envFindFlag
    clsReg.WriteString regPath, "LastKey", tvKeys.SelectedItem.FullPath
    Dim sTmp As String
    sTmp = Word2Hex(44) & Word2Hex(Me.WindowState) & _
            Word2Hex(IIf(Me.WindowState = vbMaximized, 3, 1)) & _
            "FFFFFFFF" & "FFFFFFFF" & "FFFFFFFF" & "FFFFFFFF"
    Me.Visible = False
    Me.WindowState = vbNormal
    sTmp = sTmp & _
            Word2Hex(Me.ScaleX(Int(Abs(Me.Left)), vbTwips, vbPixels)) & _
            Word2Hex(Me.ScaleY(Int(Abs(Me.Top)), vbTwips, vbPixels)) & _
            Word2Hex(Me.ScaleX(Int(Me.Width), vbTwips, vbPixels)) & _
            Word2Hex(Me.ScaleY(Int(Me.Height), vbTwips, vbPixels)) & _
            Word2Hex(Me.ScaleX(Int(Abs(pbResize.Left)), vbTwips, vbPixels))
    With lvValues
        sTmp = sTmp & Word2Hex(Me.ScaleX(.ColumnHeaders(1).Width, vbTwips, vbPixels)) & _
                    Word2Hex(Me.ScaleX(.ColumnHeaders(2).Width, vbTwips, vbPixels)) & _
                    Word2Hex(Me.ScaleX(.ColumnHeaders(3).Width, vbTwips, vbPixels))
    End With
    sTmp = sTmp & Word2Hex(IIf(mnuStatusBar.Visible, 1, 0))
    clsReg.WriteBinary regPath, "View", clsReg.BinToStr(clsReg.StrToBin(sTmp))
    
    Set clsReg = Nothing
End Sub

Private Sub lvValues_BeforeLabelEdit(Cancel As Integer)
On Error Resume Next
    If lvValues.SelectedItem.Text = sDefaultName Or lvValues.SelectedItem.Index = 1 Then Cancel = 1
End Sub

Private Sub lvValues_DblClick()
On Error GoTo exitSub
    If Not isDblClk Then Exit Sub
    With lvValues.HitTest(mX, mY)
        Select Case .ListSubItems(1).Text
            Case "REG_SZ", "REG_MULTI_SZ", "REG_EXPAND_SZ"
                If .Text <> sDefaultName Then frmEditString.txtName.Text = .Text
                If .ListSubItems(2).Text = sDefaultValue Then
                    If .Text <> sDefaultName Then
                        frmEditString.txtData.Text = sDefaultValue
                    Else
                        If .Tag <> "" Then frmEditString.txtData.Text = .ListSubItems(2).Text
                    End If
                Else
                    frmEditString.txtData.Text = .ListSubItems(2).Text
                End If
                frmEditString.txtType.Text = .ListSubItems(1).Text
                frmEditString.Show vbModal, Me
            Case "REG_BINARY"
                frmEditBIN.Show vbModal, Me
            Case "REG_DWORD"
                frmEditDWORD.txtName.Text = .Text
                frmEditDWORD.txtData.Text = Hex$(Val(Mid(.ListSubItems(2).Text, 13, Len(.ListSubItems(2).Text) - 13)))
                frmEditDWORD.Show vbModal, Me
        End Select
    End With
exitSub:
End Sub

Private Sub lvValues_GotFocus()
    curPos = "lv"
    proccessMenu
End Sub

Private Sub lvValues_KeyUp(KeyCode As Integer, Shift As Integer)
    proccessKeyCode KeyCode
End Sub

Private Sub lvValues_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    isDblClk = False
    isRMouse = Button = 2
End Sub

Private Sub lvValues_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mX = x
    mY = y
    isRMouse = False
On Error Resume Next
    mnuModify.Enabled = lvValues.SelectedItem.Selected
    mnuModifyBIN.Enabled = mnuModify.Enabled
    mnuDelete.Enabled = mnuModify.Enabled
    mnuRename.Enabled = mnuModify.Enabled
End Sub

Private Sub lvValues_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    isDblClk = Button = 1
On Error Resume Next
    If Button = 1 Then
        If Not lvValues.HitTest(x, y).Selected Then lvValues.SelectedItem.Selected = False
        GoTo exitSub
    End If
On Error GoTo popupOther
    Dim isM As Boolean
    If Button = 2 And lvValues.HitTest(x, y).Selected Then
        With lvValues.HitTest(x, y)
            isM = isMulti
            hideAllSubEditMenu
            mnuModify.Visible = True
            mnuModify.Enabled = True And Not isM
            mnuModifyBIN.Visible = True
            mnuSep00.Visible = True
            mnuDelete.Visible = True
            mnuDelete.Enabled = True
            mnuRename.Visible = True
            mnuRename.Enabled = .Text <> sDefaultName And Not isM
            mnuNew.Visible = False
            PopupMenu mnuEdit, , , , mnuModify
        End With
    End If
    GoTo exitSub
popupOther:
    openPopup x, y
exitSub:
    isRMouse = False
    proccessMenu
End Sub

Private Sub mnuAbout_Click()
    ShellAbout Me.hWnd, Me.Caption, "Copyright © 2006 by sowLov - BaCuong", Me.Icon
End Sub

Private Sub mnuAddFavorites_Click()
    frmAddFav.txtName.Text = tvKeys.SelectedItem.Text
    frmAddFav.lblValue.Caption = tvKeys.SelectedItem.FullPath
    frmAddFav.Show vbModal, Me
End Sub

Private Sub mnuCollapse_Click()
On Error Resume Next
    tvKeys.SelectedItem.Expanded = False
End Sub

Private Sub mnuCopy_Click()
On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText explainPath(tvKeys.SelectedItem.FullPath)
End Sub

Private Sub mnuDelete_Click()
On Error GoTo exitSub
    If curPos = "tv" Then
        If MsgBox("Are you sure you want to delete this key and all of its subkeys?", vbExclamation + vbYesNo, "Confirm Key Delete") = vbYes Then
            clsReg.KillKey explainPath(tvKeys.SelectedItem.FullPath)
            tvKeys.Nodes.Remove tvKeys.SelectedItem.Index
            addnewValues tvKeys.SelectedItem
        End If
    Else
        If MsgBox("Are you sure you want to delete " & IIf(isMulti(), "these values", "this value") & "?", vbExclamation + vbYesNo, "Confirm Value Delete") = vbYes Then
            Dim selItem As Integer
            For selItem = lvValues.ListItems.Count To 1 Step -1
                If lvValues.ListItems(selItem).Selected Then
                    If lvValues.ListItems(selItem).Text = sDefaultName Then
                        If Not clsReg.ValueExists(explainPath(sbStastus.SimpleText), "") Then
                            MsgBox "Unabled to delete all specified values.", vbCritical, "Error Deleting Values"
                        Else
                            clsReg.KillValue explainPath(sbStastus.SimpleText), ""
                            lvValues.ListItems(selItem).ListSubItems(2).Text = sDefaultValue
                        End If
                    Else
                        clsReg.KillValue explainPath(sbStastus.SimpleText), lvValues.ListItems(selItem).Text
                        lvValues.ListItems.Remove lvValues.ListItems(selItem).Index
                    End If
                End If
            Next selItem
        End If
    End If
exitSub:
End Sub

Private Sub mnuEditExport_Click()
    mnuExport_Click
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExpand_Click()
On Error Resume Next
    tvKeys.SelectedItem.Expanded = True
End Sub

Private Sub mnuExport_Click()
    With cdDialog
        .DialogTitle = "Export Registry File"
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
        If .FileName <> "" Then
            If Dir(.FileName) <> "" Then Kill .FileName
            clsReg.ExportToReg .FileName, sbStastus.SimpleText, True
        End If
    End With
End Sub

Private Sub mnuFavIdx_Click(Index As Integer)
    If Index > 0 Then gotoAKey mnuFavIdx(Index).Tag
End Sub

Private Sub mnuFind_Click()
    frmFind.Show vbModal, Me
    If ((envFindFlag Or 8) = 8) Or ((envFindFlag Or 4) = 4) Then lvValues.SetFocus
End Sub

Private Sub mnuFindNext_Click()
    If fValue <> "" Then frmFinding.Show vbModal, Me Else frmFind.Show vbModal, Me
End Sub

Private Sub mnuFindTree_Click()
    mnuFindNext_Click
End Sub

Private Sub mnuImport_Click()
    With cdDialog
        .DialogTitle = "Import Registry File"
        .Flags = cdlOFNOverwritePrompt
        .ShowOpen
        If .FileName <> "" Then
            If Dir(.FileName) = "" Then Exit Sub
            clsReg.ImportFromReg .FileName
        End If
    End With
End Sub

Private Sub mnuModifyBIN_Click()
    isDblClk = True: lvValues_DblClick: isDblClk = False
End Sub

Private Sub mnuNewExpString_Click()
    addnewValue REG_EXPAND_SZ
End Sub

Private Sub mnuModify_Click()
    isDblClk = True: lvValues_DblClick: isDblClk = False
End Sub

Private Sub mnuNewBinary_Click()
    addnewValue REG_BINARY
End Sub

Private Sub mnuNewDWORD_Click()
    addnewValue REG_DWORD
End Sub

Private Sub mnuNewKey_Click()
On Error GoTo notExist
    Dim nCount As Integer, selItem As MSComctlLib.Node, newNode As MSComctlLib.Node
    nCount = 1
    Set selItem = tvKeys.SelectedItem
    While clsReg.KeyExists(explainPath(selItem.FullPath) & "\New Key #" & nCount)
        nCount = nCount + 1
    Wend
    If clsReg.CreateKey(explainPath(selItem.FullPath) & "\New Key #" & nCount) = 0 Then
        MsgBox "Cannot create key: Error while opening the key " & selItem.Text & ".", vbCritical, "Error Creating Key"
    Else
        Set newNode = tvKeys.Nodes.Add(selItem, tvwChild, , "New Key #" & nCount, "Closed", "Opened")
        addnewValues newNode
        newNode.Selected = True
        tvKeys.StartLabelEdit
    End If
notExist:
    Set selItem = Nothing
    Set newNode = Nothing
End Sub

Private Sub mnuNewMultiString_Click()
    addnewValue REG_MULTI_SZ
End Sub

Private Sub mnuNewString_Click()
    addnewValue REG_SZ
End Sub

Private Sub mnuOption_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuRefresh_Click()
    addnewValues tvKeys.SelectedItem
End Sub

Private Sub mnuRemoveFavorite_Click()
    frmGetFav.Show vbModal, Me
End Sub

Private Sub mnuRename_Click()
On Error GoTo exitSub
    With IIf(curPos = "tv", tvKeys, lvValues)
        .SelectedItem.Selected = True
        .StartLabelEdit
    End With
exitSub:
End Sub

Private Sub mnuStatusBar_Click()
    mnuStatusBar.Checked = Not mnuStatusBar.Checked
    sbStastus.Visible = mnuStatusBar.Checked
    Form_Resize
End Sub

Private Sub mnuTools_Click(Index As Integer)
    If Index > 0 Then gotoAKey tvKeys.Nodes(1).Text & "\" & mnuTools(Index).Tag
End Sub

Private Sub pbResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pbResize.BackColor = &H808080
    isDrag = True
End Sub

Private Sub pbResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x < 0 And Me.ScaleX(pbResize.Left, vbTwips, vbPixels) < 100 Then Exit Sub
    If x > 0 And Me.ScaleX(pbResize.Left, vbTwips, vbPixels) > Me.ScaleX(Me.ScaleWidth, vbTwips, vbPixels) - 100 Then Exit Sub
    If isDrag Then pbResize.Left = pbResize.Left + x
End Sub

Private Sub pbResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    isDrag = False
    pbResize.BackColor = vbButtonFace
    Form_Resize
End Sub

Private Sub tvKeys_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error GoTo exitSub
    Dim newKey As String
    newKey = explainPath(oldName.FullPath)
    newKey = Mid(newKey, 1, InStrRev(newKey, "\")) & NewString
    If clsReg.KeyExists(newKey) Then
        MsgBox "The Registry Editor cannot rename " & oldName.Text & ". The specified key name already exists. Type another name and try again.", vbCritical, "Error Renaming Key"
        Cancel = 1
    Else
        If Not clsReg.renameKey(explainPath(oldName.FullPath), NewString) Then
            MsgBox "The Registry Editor cannot rename " & oldName.Text & ". Error while renaming key.", vbCritical, "Error Renaming Key"
            Cancel = 1
        End If
    End If
    Set oldName = Nothing
exitSub:
End Sub

Private Sub tvKeys_BeforeLabelEdit(Cancel As Integer)
On Error GoTo exitSub
    Set oldName = tvKeys.SelectedItem
exitSub:
End Sub

Private Sub tvKeys_Expand(ByVal Node As MSComctlLib.Node)
On Error GoTo exitSub
    Dim nChild As Integer, curSub As MSComctlLib.Node
    If Node.Children > 0 And explainPath(Node.FullPath) <> "" Then
        sbStastus.SimpleText = "Loading..."
        Set curSub = Node.Child
        addnewSubKeys curSub
        For nChild = 1 To Node.Children - 1
            Set curSub = curSub.Next
            addnewSubKeys curSub
        Next nChild
        sbStastus.SimpleText = Node.FullPath
    End If
exitSub:
    Set curSub = Nothing
End Sub

Private Sub tvKeys_GotFocus()
    curPos = "tv"
    proccessMenu
End Sub

Private Sub tvKeys_KeyUp(KeyCode As Integer, Shift As Integer)
    proccessKeyCode KeyCode
    If KeyCode = 13 Then sbStastus.SimpleText = tvKeys.SelectedItem.FullPath
End Sub

Private Sub tvKeys_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    isRMouse = Button = 2
End Sub

Private Sub tvKeys_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mX = x
    mY = y
    isRMouse = False
End Sub

Private Sub tvKeys_NodeClick(ByVal Node As MSComctlLib.Node)
    sbStastus.SimpleText = "Loading..."
    addnewSubKeys Node
    addnewValues Node
    proccessMenu
    If isRMouse Then openPopup x, y: isRMouse = False
End Sub
