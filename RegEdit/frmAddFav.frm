VERSION 5.00
Begin VB.Form frmAddFav 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add to Favorites"
   ClientHeight    =   930
   ClientLeft      =   1050
   ClientTop       =   1335
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddFav.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3240
      TabIndex        =   3
      Top             =   495
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   3240
      TabIndex        =   2
      Top             =   75
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   3060
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value..."
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   690
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Favorite name:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    With frmRegEdit
        If .clsReg.ValueExists(.regPath & "\Favorites", txtName.Text) Then
            MsgBox "There is already a favorite with that name.", vbCritical, "Error Adding Favorite"
            Exit Sub
        Else
            .clsReg.WriteString .regPath & "\Favorites", txtName.Text, lblValue.Caption
        End If
        .getFavorites
    End With
    Unload Me
End Sub
