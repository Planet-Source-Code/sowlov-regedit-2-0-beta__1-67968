VERSION 5.00
Begin VB.Form frmFinding 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1620
   ClientLeft      =   1050
   ClientTop       =   1335
   ClientWidth     =   3915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinding.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   Begin VB.Timer timSearching 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   300
      Top             =   270
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2625
      TabIndex        =   1
      Top             =   1065
      Width           =   1095
   End
   Begin VB.Label lblinKEY 
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   825
      TabIndex        =   2
      Top             =   615
      Width           =   3015
   End
   Begin VB.Label lblSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Searching the registry..."
      Height          =   195
      Left            =   825
      TabIndex        =   0
      Top             =   375
      Width           =   1770
   End
   Begin VB.Image imgSearch 
      Height          =   480
      Left            =   270
      Top             =   255
      Width           =   480
   End
End
Attribute VB_Name = "frmFinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fPath As String
Dim fValue As String
Dim fIn(3) As Boolean

Dim isFound As Boolean
Dim regFind As clsRegistryAccess
Dim maxNode As Integer
Dim curNode As Integer
Dim isScaning As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Find key(s) in path
Private Function findKey(ByVal sPath As String) As Boolean
    Dim vKeys() As String
    Dim nScan As Long, nItems As Long, lFound As Boolean
    'check simple but quick
    'compare for itself
    lFound = InStr(1, Mid(sPath, InStrRev(sPath, "\") + 1), fValue, vbTextCompare) > 0
    'check for exactly of subkey with name=fvalue
    If Not lFound Then
        lFound = regFind.KeyExists(sPath & "\" & fValue)
        If lFound Then
            fPath = sPath & "\" & fValue
        Else 'if still not found
            'Search in subkeys
            nItems = regFind.EnumKeys(sPath, vKeys)
            If nItems > 0 Then
                If Not fIn(3) Then
                    nScan = 0
                    While nScan < nItems And Not lFound
                        lFound = InStr(1, vKeys(nScan), fValue, vbTextCompare) > 0
                        If lFound Then fValue = vKeys(nScan)
                        nScan = nScan + 1
                    Wend
                End If
            End If
        End If
    End If
    findKey = lFound
End Function

'Find value(s) AND/OR in path
Private Function findValueData(ByVal sPath As String) As Boolean
    Dim vNames() As String, vDatas() As Variant
    Dim nScan As Long, nItems As Long, lFound As Boolean
    nItems = regFind.EnumValues(sPath, vNames, vDatas)
    lFound = False
    If nItems > 0 Then
        nScan = 0
        While nScan < nItems And Not lFound
            If fIn(1) Then 'checked in Values
                If fIn(3) Then
                    lFound = regFind.ValueExists(sPath, vNames(nScan))
                Else
                    lFound = InStr(1, vNames(nScan), fValue, vbTextCompare) > 0
                End If
            End If
            If fIn(2) And Not lFound Then 'checked in Data
                lFound = lFound Or InStr(1, "" & vDatas(nScan), fValue, vbTextCompare) > 0
                If fIn(3) And lFound Then lFound = (Len("" & vDatas(nScan)) = Len(fValue))
                If lFound Then fValue = vNames(nScan)
            End If
            nScan = nScan + 1
        Wend
    End If
    findValueData = lFound
End Function

Private Function nextKey(ByVal sPath As String, Optional ByVal curKey As String = "") As String
    Dim sNext As String
    Dim arrKeys() As String, nKeys As Long
    Dim nK As Long, isNext As Boolean
    sNext = ""
    isNext = False
    nKeys = regFind.EnumKeys(sPath, arrKeys)
    If nKeys = 0 Then
        curKey = Mid(sPath, InStrRev(sPath, "\") + 1)
        sPath = Mid(sPath, 1, InStrRev(sPath, "\") - 1)
        'is ROOT NODE
        If InStr(1, sPath, "\") = 0 Then
            sNext = nextROOT()
        Else
            sNext = nextKey(sPath, curKey)
        End If
    ElseIf nKeys > 0 Then
        If curKey = "" Then
            sNext = sPath & "\" & arrKeys(0)
        Else
            nK = 0
            While nK < nKeys And Not isNext
                If arrKeys(nK) = curKey Then isNext = True
                If Not isNext Then nK = nK + 1
            Wend
            If isNext And nK < nKeys - 1 Then
                sNext = sPath & "\" & arrKeys(nK + 1)
            Else
                curKey = Mid(sPath, InStrRev(sPath, "\") + 1)
                sPath = Mid(sPath, 1, InStrRev(sPath, "\") - 1)
                'is ROOT NODE
                If InStr(1, sPath, "\") = 0 Then
                    sNext = nextROOT()
                Else
                    sNext = nextKey(sPath, curKey)
                End If
            End If
        End If
    End If
    nextKey = sNext
End Function

Private Function nextROOT() As String
    Dim sRoot As String
    Dim Subs() As String, nSubs As Long
    sRoot = ""
    If curNode <= maxNode Then curNode = curNode + 1
    If curNode < maxNode Then
        sRoot = frmRegEdit.tvKeys.Nodes(curNode).Text
        nSubs = regFind.EnumKeys(sRoot, Subs)
        If nSubs > 0 Then sRoot = sRoot & "\" & Subs(0) Else sRoot = ""
    End If
    nextROOT = sRoot
End Function

Private Sub scanKeys()
    If Not isFound Then
        If fIn(0) Then isFound = findKey(fPath)
        If Not isFound And (fIn(1) Or fIn(2)) Then isFound = findValueData(fPath)
        If Not isFound Then fPath = nextKey(fPath)
    End If
    isScaning = False
End Sub

Private Sub Form_Load()
    Dim sTmp As String
    Set imgSearch.Picture = frmRegEdit.ilImages.ListImages("Finding").Picture
    With frmRegEdit
        fPath = Replace(.tvKeys.SelectedItem.FullPath, .tvKeys.Nodes(1).Text & "\", "")
        fValue = .fValue
        fIn(0) = (.envFindFlag And 2) = 2
        fIn(1) = (.envFindFlag And 4) = 4
        fIn(2) = (.envFindFlag And 8) = 8
        fIn(3) = (.envFindFlag And 1) = 1
    End With
    Unload frmFind
    isFound = False
    curNode = 0
    sTmp = Left(fPath, InStr(1, fPath & "\", "\") - 1)
    With frmRegEdit.tvKeys
        maxNode = .Nodes("\").Children
        Do While curNode < .Nodes("\").Children
            curNode = curNode + 1
            If .Nodes(.Nodes("\").Index + curNode).Text = sTmp Then curNode = curNode + 1: Exit Do
        Loop
    End With
    Set regFind = New clsRegistryAccess
    isScaning = False
    If Not fIn(1) Or Not fIn(2) Then fPath = nextKey(fPath)
    timSearching.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set regFind = Nothing
End Sub

Private Sub timSearching_Timer()
    If Not isScaning Then
        isScaning = True
        scanKeys
    End If
    If isFound Or curNode > maxNode Then
        timSearching.Enabled = False
        If isFound Then
            Dim nLI As Integer
            frmRegEdit.gotoAKey frmRegEdit.tvKeys.Nodes("\").Text & "\" & fPath & "\" & fValue
            If fIn(1) Or fIn(2) Then
                With frmRegEdit.lvValues
                    For nLI = 1 To .ListItems.Count
                        .ListItems(nLI).Selected = UCase(.ListItems(nLI).Text) = UCase(fValue)
                    Next nLI
                End With
            End If
        End If
        If curNode > maxNode Then MsgBox "Finished searching throught the registry.", vbQuestion, frmRegEdit.Caption
        Unload Me
    End If
End Sub
