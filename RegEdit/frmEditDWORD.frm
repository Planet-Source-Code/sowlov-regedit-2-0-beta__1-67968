VERSION 5.00
Begin VB.Form frmEditDWORD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit DWORD Value"
   ClientHeight    =   2235
   ClientLeft      =   1050
   ClientTop       =   1335
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditDWORD.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   316
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2340
      TabIndex        =   8
      Top             =   1770
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3540
      TabIndex        =   7
      Top             =   1770
      Width           =   1095
   End
   Begin VB.Frame frmBase 
      Caption         =   "Base"
      Height          =   990
      Left            =   2355
      TabIndex        =   4
      Top             =   690
      Width           =   2280
      Begin VB.OptionButton optNum 
         Caption         =   "Decimal"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   615
         Width           =   1800
      End
      Begin VB.OptionButton optNum 
         Caption         =   "Hexadecimal"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1800
      End
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   4545
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Left            =   90
      MaxLength       =   7
      TabIndex        =   3
      Top             =   930
      Width           =   1755
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value name:"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   60
      Width           =   885
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value data:"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   690
      Width           =   825
   End
End
Attribute VB_Name = "frmEditDWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim nVal As Double
    With frmRegEdit
        If optNum(0).Value Then
            nVal = .Hex2Dbl(txtData.Text)
        Else
            nVal = Val(txtData.Text)
        End If
        .clsReg.WriteDWORD .explainPath(.tvKeys.SelectedItem.FullPath), txtName.Text, nVal
        .addnewValues .tvKeys.SelectedItem
    End With
    Unload Me
End Sub

Private Sub Form_Activate()
    optNum_Click (0)
    txtData.SelStart = 0
    txtData.SelLength = Len(txtData.Text)
    txtData.SetFocus
End Sub

Private Sub optNum_Click(Index As Integer)
    With txtData
        If Index = 0 Then '-> HEX
            If .MaxLength = 9 Then
                .Text = frmRegEdit.Dbl2Hex(Val(.Text))
                .MaxLength = 7
            End If
        Else '-> DEC
            If .MaxLength = 7 Then
                .MaxLength = 9
                .Text = frmRegEdit.Hex2Dbl(.Text)
            End If
        End If
    End With
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    Const strHEX = "0123456789ABCDEF"
    Dim cChar As String
    If KeyAscii <> 8 Then
        cChar = UCase(Chr$(KeyAscii))
        If optNum(0).Value = True Then
            KeyAscii = IIf(InStr(1, strHEX, cChar) = 0, 0, Asc(cChar))
        Else
            If InStr(1, Mid(strHEX, 1, 10), cChar) = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtData_KeyUp(KeyCode As Integer, Shift As Integer)
    Const maxVal = 268435455
    If Val(txtData.Text) > maxVal Then txtData.Text = maxVal: txtData.SelStart = 0: txtData.SelLength = Len(txtData.Text)
End Sub
