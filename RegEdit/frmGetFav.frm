VERSION 5.00
Begin VB.Form frmGetFav 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remove Favorites"
   ClientHeight    =   2325
   ClientLeft      =   1050
   ClientTop       =   1335
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetFav.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstFavs 
      Height          =   1815
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   390
      Width           =   3195
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3405
      TabIndex        =   3
      Top             =   825
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   3405
      TabIndex        =   2
      Top             =   390
      Width           =   1095
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Favorite(s):"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1335
   End
End
Attribute VB_Name = "frmGetFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim nI As Integer
    With frmRegEdit
        For nI = 0 To lstFavs.ListCount - 1
            If lstFavs.Selected(nI) Then .clsReg.KillValue .regPath & "\Favorites", Trim(lstFavs.List(nI))
        Next nI
        .getFavorites
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sNames() As String, sValues() As Variant, nIdx As Long
    Dim I As Long
    With frmRegEdit
        nIdx = .clsReg.EnumValues(.regPath & "\Favorites", sNames, sValues, REG_SZ)
        For I = 0 To nIdx - 1
            lstFavs.AddItem sNames(I)
        Next I
    End With
End Sub
