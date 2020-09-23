VERSION 5.00
Begin VB.Form frmBombSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Bomb"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1935
   ControlBox      =   0   'False
   DrawMode        =   3  'Not Merge Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   1935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.OptionButton optMega 
      Caption         =   "Artillery"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton optMax 
      Caption         =   "Atomic Bomb"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.OptionButton optLarge 
      Caption         =   "Anti-Aircraft Cannon"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.OptionButton optMedium 
      Caption         =   "Bazooka"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton optSmall 
      Caption         =   "Fireworks"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmBombSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If optSmall.Value Then
Tag = 0.5
ElseIf optMedium.Value Then
Tag = 2.5
ElseIf optLarge.Value Then
Tag = 5.5
ElseIf optMega.Value Then
Tag = 15.5
ElseIf optMax.Value Then
Tag = 40
End If
Hide
End Sub

