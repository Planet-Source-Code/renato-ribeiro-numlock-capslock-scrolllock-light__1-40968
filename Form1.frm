VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Light on computer"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   2955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Powered by:                                                Renato Ribeiro"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Renato Ribeiro
'Portugal
'
'prenatoribeiro@sapo.pt
'
'a funny program
'
Private Sub Form_Load()
    vk_numlock = &H90 '144
    vk_capslock = &H14 '20
    vk_scrolllock = &H91 '145
    light = 1
End Sub

Private Sub Form_Terminate()
    lightOff
End Sub

Private Sub Timer1_Timer()
    lightOn
End Sub

