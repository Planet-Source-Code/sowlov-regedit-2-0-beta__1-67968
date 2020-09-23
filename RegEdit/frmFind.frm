VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2325
   ClientLeft      =   1050
   ClientTop       =   1335
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4755
      TabIndex        =   8
      Top             =   525
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   360
      Left            =   4755
      TabIndex        =   7
      Top             =   90
      Width           =   1095
   End
   Begin VB.CheckBox chkMatch 
      Caption         =   "Match &whole string only"
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4440
   End
   Begin VB.Frame frmLook 
      Caption         =   " Look at "
      Height          =   1260
      Left            =   120
      TabIndex        =   2
      Top             =   555
      Width           =   4530
      Begin VB.CheckBox chkData 
         Caption         =   "&Data"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   4305
      End
      Begin VB.CheckBox chkValues 
         Caption         =   "&Value"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   615
         Width           =   4305
      End
      Begin VB.CheckBox chkKey 
         Caption         =   "&Keys"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   285
         Width           =   4305
      End
   End
   Begin VB.TextBox txtFindString 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   105
      Width           =   3735
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fi&nd what:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub canFind()
    cmdFind.Enabled = txtFindString.Text <> "" And _
        (chkKey.Value = vbChecked Or chkValues.Value = vbChecked Or chkData.Value = vbChecked)
End Sub

Private Sub chkData_Click()
    canFind
End Sub

Private Sub chkKey_Click()
    canFind
End Sub

Private Sub chkValues_Click()
    canFind
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    frmRegEdit.envFindFlag = _
        IIf(chkData.Value = vbChecked, 8, 0) Xor _
        IIf(chkValues.Value = vbChecked, 4, 0) Xor _
        IIf(chkKey.Value = vbChecked, 2, 0) Xor _
        IIf(chkMatch.Value = vbChecked, 1, 0)
    frmRegEdit.fValue = txtFindString.Text
    frmFinding.Show vbModal, frmRegEdit
End Sub

Private Sub Form_Load()
    With frmRegEdit
        chkData.Value = IIf((.envFindFlag And 8) = 8, vbChecked, vbUnchecked)
        chkValues.Value = IIf((.envFindFlag And 4) = 4, vbChecked, vbUnchecked)
        chkKey.Value = IIf((.envFindFlag And 2) = 2, vbChecked, vbUnchecked)
        chkMatch.Value = IIf((.envFindFlag And 1) = 1, vbChecked, vbUnchecked)
        txtFindString.Text = .fValue
    End With
    canFind
End Sub

Private Sub txtFindString_Change()
    canFind
End Sub
