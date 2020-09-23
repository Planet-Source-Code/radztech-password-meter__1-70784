VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Meter"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1110
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   810
      Width           =   2235
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   300
      Width           =   2265
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   1110
      TabIndex        =   4
      Top             =   1290
      Width           =   2265
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   870
      Width           =   735
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text2_Change()
Label1.Caption = PWMeter(Text1.Text, Text2.Text)
End Sub

