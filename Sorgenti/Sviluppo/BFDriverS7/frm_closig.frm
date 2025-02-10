VERSION 5.00
Begin VB.Form frm_closing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BFDriver_S7"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_chiudi 
      Caption         =   "Chiudi"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmd_nascondi 
      Caption         =   "Nascondi"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frm_closing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_chiudi_Click()

  close_cmd = True
  Unload Me
  
End Sub

Private Sub cmd_nascondi_Click()

  hide_cmd = True
  Unload Me
  
End Sub

Private Sub Form_Load()

  close_cmd = False
  hide_cmd = False
  
End Sub
