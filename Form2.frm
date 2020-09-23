VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Choose color"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form2"
   ScaleHeight     =   1995
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtColor 
      Height          =   375
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "12"
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblCuloare 
      Caption         =   "Choose color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  c = CInt(txtColor.Text)
If (c > 0 And c < 16) Then
 Unload Me
Else
 c = 12
 MsgBox "Value is not correct !", , "Choose color..."
 Unload Me
End If
End Sub


