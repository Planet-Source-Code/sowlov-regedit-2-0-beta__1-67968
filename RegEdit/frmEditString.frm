VERSION 5.00
Begin VB.Form frmEditString 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit String"
   ClientHeight    =   1695
   ClientLeft      =   1050
   ClientTop       =   1335
   ClientWidth     =   5310
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmEditString.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   354
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDataMulti 
      Height          =   300
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   915
      Visible         =   0   'False
      Width           =   5145
   End
   Begin VB.TextBox txtType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5010
      TabIndex        =   6
      Top             =   615
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4125
      TabIndex        =   5
      Top             =   1290
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2925
      TabIndex        =   4
      Top             =   1290
      Width           =   1095
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Left            =   75
      TabIndex        =   3
      Top             =   915
      Width           =   5145
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   285
      Width           =   5145
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value data:"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   675
      Width           =   825
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value name:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   885
   End
End
Attribute VB_Name = "frmEditString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim sPath As String
    With frmRegEdit
        sPath = .explainPath(.sbStastus.SimpleText)
        Select Case txtType.Text
            Case "REG_EXPAND_SZ": .clsReg.WriteString sPath, txtName.Text, txtData.Text, REG_EXPAND_SZ
            Case "REG_MULTI_SZ": .clsReg.WriteString sPath, txtName.Text, txtDataMulti.Text, REG_MULTI_SZ
            Case Else: .clsReg.WriteString sPath, txtName.Text, txtData.Text
        End Select
        .lvValues.SelectedItem.ListSubItems(2).Text = IIf(txtType.Text = "REG_MULTI_SZ", txtDataMulti.Text, txtData.Text)
    End With
    Unload Me
End Sub

Private Sub Form_Activate()
    If txtType.Text = "REG_MULTI_SZ" Then
        With txtDataMulti
            Const expLine = 7
            cmdOK.Top = (cmdOK.Top - .Height) + (.Height * expLine)
            cmdCancel.Top = cmdOK.Top
            Me.Height = Me.ScaleY((Me.ScaleHeight - .Height + cmdOK.Height) + (.Height * expLine), vbPixels, vbTwips)
            .Height = .Height * expLine
            txtData.Visible = False
            .Visible = True
            .Text = txtData.Text
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    Else
        With txtData
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
End Sub
