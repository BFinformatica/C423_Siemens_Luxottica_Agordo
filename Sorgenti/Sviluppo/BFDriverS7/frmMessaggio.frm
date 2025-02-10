VERSION 5.00
Begin VB.Form frmMessaggio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Protocollo S7"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.Label msg 
      Alignment       =   2  'Center
      Caption         =   "Messaggio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "frmMessaggio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
