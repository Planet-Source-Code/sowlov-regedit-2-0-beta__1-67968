VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options Tool"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   501
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   594
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtKey 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   6660
      Width           =   2820
   End
   Begin VB.TextBox txtData 
      Height          =   330
      Left            =   2970
      TabIndex        =   5
      Top             =   6660
      Width           =   5880
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   360
      Left            =   6600
      TabIndex        =   0
      Top             =   7095
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   7755
      TabIndex        =   2
      Top             =   7095
      Width           =   1095
   End
   Begin MSComctlLib.TreeView tvTrees 
      Height          =   6090
      Left            =   75
      TabIndex        =   1
      Top             =   270
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10742
      _Version        =   393217
      Indentation     =   132
      Style           =   7
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of Windows invisible resource:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   30
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key Value:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2985
      TabIndex        =   4
      Top             =   6405
      Width           =   870
   End
   Begin VB.Label lblKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   6405
      Width           =   885
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
    Dim chdIndex As Integer
    Dim sPath As String, sName As String
    If Node.Image = "item" Then
        chdIndex = CInt(Node.Tag)
        sPath = arrChd(chdIndex, 6)
        If arrChd(chdIndex, 6) = "" Then sPath = arrPar(arrChd(chdIndex, 1), 3)
        sName = arrChd(chdIndex, 3)
        Select Case arrChd(chdIndex, 4)
            Case v_oneZERO: frmRegEdit.clsReg.WriteDWORD sPath, sName, Val("0" & txtData.Text)
            Case v_String: frmRegEdit.clsReg.WriteString sPath, sName, Str(txtData.Text)
            Case v_Binary: frmRegEdit.clsReg.WriteBinary sPath, sName, frmRegEdit.clsReg.StrToBin(txtData.Text)
        End Select
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Set tvTrees.ImageList = frmRegEdit.ilImages
    Dim clsReg As clsRegistryAccess
    Set clsReg = New clsRegistryAccess
    
    initVarible
    loadData
    
    Dim nTree As Integer
    For nTree = 1 To maxPar
        If arrPar(nTree, 1) = 0 Then
            With tvTrees.Nodes.Add(, tvwRootLines, "KEY" & nTree, arrPar(nTree, 2), "Closed", "Opened")
                .Tag = arrPar(nTree, 3)
                .Sorted = True
                .Expanded = True
            End With
        Else
            With tvTrees.Nodes.Add("KEY" & arrPar(nTree, 1), tvwChild, "KEY" & nTree, arrPar(nTree, 2), "Closed", "Opened")
                .Tag = arrPar(nTree, 3)
                .Sorted = True
                .Expanded = True
            End With
        End If
    Next nTree
    For nTree = 1 To maxChd
        With tvTrees.Nodes.Add("KEY" & arrChd(nTree, 1), tvwChild, , arrChd(nTree, 2), "item", "item")
            .Tag = nTree
        End With
    Next nTree
    Set clsReg = Nothing
End Sub

Private Sub tvTrees_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim chdIndex As Integer
    If Node.Image = "item" Then
        chdIndex = CInt(Node.Tag)
        txtKey.Text = arrChd(chdIndex, 3)
        txtData.Text = arrChd(chdIndex, 5)
    Else
        txtKey.Text = ""
        txtData.Text = ""
    End If
    cmdApply.Enabled = Node.Image = "item"
End Sub
