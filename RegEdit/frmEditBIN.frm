VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEditBIN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Binary Value"
   ClientHeight    =   4335
   ClientLeft      =   1050
   ClientTop       =   1335
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditBIN.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5055
      TabIndex        =   10
      Top             =   690
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   5175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4170
      TabIndex        =   5
      Top             =   3900
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2955
      TabIndex        =   4
      Top             =   3900
      Width           =   1095
   End
   Begin VB.PictureBox picRTFBase 
      BackColor       =   &H80000009&
      Height          =   2790
      Left            =   90
      ScaleHeight     =   182
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   341
      TabIndex        =   0
      Top             =   990
      Width           =   5175
      Begin RichTextLib.RichTextBox RTFLine 
         Height          =   2730
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   4815
         _Version        =   393217
         BorderStyle     =   0
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmEditBIN.frx":014A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTFAscii 
         Height          =   2730
         Left            =   3710
         TabIndex        =   1
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   4815
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmEditBIN.frx":01FB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTFHex 
         Height          =   2730
         Left            =   735
         TabIndex        =   2
         Top             =   0
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   4815
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmEditBIN.frx":02E4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSForms.ScrollBar sbScroll 
         Height          =   2730
         Left            =   4860
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         ForeColor       =   16711680
         Size            =   "450;4815"
      End
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   1020
      TabIndex        =   9
      Top             =   735
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value data :"
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   735
      Width           =   870
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value name :"
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   120
      Width           =   930
   End
   Begin VB.Menu mnuEditBase 
      Caption         =   "Edit Base"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Cut"
         Index           =   0
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Delete"
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select All"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmEditBIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

Const lineHeight = 13
Const RowsPerPage = 14
Dim arrBIN As Variant

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub setupScrollBar()
    Dim nRows As Double ', nI As Long
    nRows = Len(RTFAscii.Text) \ 9
    If Len(RTFAscii.Text) Mod 9 > 0 Then nRows = nRows + 1
    LockWindowUpdate Me.hWnd
'    RTFLine.Text = ""
'    For nI = 0 To nRows
'        RTFLine.Text = RTFLine.Text & frmRegEdit.Dbl2Hex(8 * nI, 4)
'        If nI < nRows Then RTFLine.Text = RTFLine.Text & Chr(13)
'    Next
    If nRows < RowsPerPage Then
        sbScroll.Visible = False
        RTFLine.Top = 0
        RTFLine.Height = RowsPerPage * lineHeight
        RTFHex.Top = 0
        RTFHex.Height = RowsPerPage * lineHeight
        RTFAscii.Top = 0
        RTFAscii.Height = RowsPerPage * lineHeight
    Else
        nRows = nRows + 1
        sbScroll.Visible = True
        sbScroll.Max = nRows - RowsPerPage
        RTFLine.Height = nRows * lineHeight
        RTFHex.Height = nRows * lineHeight
        RTFAscii.Height = nRows * lineHeight
    End If
    LockWindowUpdate 0
End Sub

Private Sub Form_Load()
    frmLoading.Show vbModeless, Me
    RTFLine.Text = "0000"
    RTFHex.Text = ""
    RTFAscii.Text = ""
    With frmRegEdit.lvValues.SelectedItem
        txtName.Text = .Text
        txtType.Text = .ListSubItems(1).Text
        lblPath.Caption = frmRegEdit.explainPath(frmRegEdit.tvKeys.SelectedItem.FullPath)
    End With
    Dim nBINs As Long
    Dim nInt As Long, curVal As Byte
    Dim sA As String, sB As String, sC As String
    arrBIN = frmRegEdit.clsReg.ReadBinary(lblPath.Caption, txtName.Text, , BIN_Array)
    If VarType(arrBIN) = vbArray + vbByte Then
        nBINs = UBound(arrBIN)
        For nInt = 0 To nBINs - 1
            LockWindowUpdate Me.hWnd
            curVal = arrBIN(nInt)
            If nInt > 0 Then
                If nInt Mod 8 = 0 Then
                    RTFLine.Text = RTFLine.Text & Chr$(13) & frmRegEdit.Dbl2Hex(8 * (nInt \ 8), 4)
                    RTFHex.Text = RTFHex.Text & Chr$(13)
                    RTFAscii.Text = RTFAscii.Text & Chr$(13)
                Else
                    RTFHex.Text = RTFHex.Text & Chr$(32)
                End If
            End If
            RTFHex.Text = RTFHex.Text & IIf(curVal < &HF, Chr$(48), "") & Hex$(curVal)
            If curVal < 33 Or (curVal > 126 And curVal < 144) Or (curVal > 147 And curVal < 161) Then
                RTFAscii.Text = RTFAscii.Text & Chr(46)
            Else
                RTFAscii.Text = RTFAscii.Text & Chr$(curVal)
            End If
            LockWindowUpdate 0
        Next nInt
    End If
    setupScrollBar
    Unload frmLoading
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase arrBIN
    RTFLine.Text = ""
    RTFHex.Text = ""
    RTFAscii.Text = ""
End Sub

Private Sub RTFAscii_Change()
'    RTFHex.SelStart = RTFLine.GetLineFromChar(RTFLine.SelStart) * 24
End Sub

Private Sub RTFAscii_SelChange()
    Dim nL As Integer, nC As Integer
    nL = RTFAscii.SelStart \ 9
    nC = RTFAscii.SelStart Mod 8
    RTFHex.SelStart = (nL * 24) + (nC * 3)
    RTFHex.SelLength = RTFAscii.SelLength * 2
End Sub

Private Sub RTFHex_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And _
        (KeyAscii < 97 Or KeyAscii > 102) And _
        (KeyAscii < 65 Or KeyAscii > 70) Then KeyAscii = 0
End Sub

Private Sub RTFLine_GotFocus()
    RTFHex.SelStart = RTFLine.GetLineFromChar(RTFLine.SelStart) * 24
    RTFHex.SetFocus
End Sub

Private Sub RTFLine_SelChange()
    RTFLine.SelLength = 0
End Sub

Private Sub sbScroll_Change()
    RTFLine.Top = sbScroll.Top - (sbScroll.Value * lineHeight)
    RTFHex.Top = sbScroll.Top - (sbScroll.Value * lineHeight)
    RTFAscii.Top = sbScroll.Top - (sbScroll.Value * lineHeight)
End Sub
