VERSION 5.00
Begin VB.Form faq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help FAQ"
   ClientHeight    =   1500
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "DOCX Format ( MSO Word )"
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Windows Help Format ( Need to be Installed )"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3972
   End
End
Attribute VB_Name = "faq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "Error Opening File /Help/SUPPORT.hlp! Please, Open it Manually.", vbOKOnly + vbCritical, "File Access Error!"
End Sub

Private Sub Command2_Click()
MsgBox "Error Opening File /Help/doc.docx! Please, Open it Manually.", vbOKOnly + vbCritical, "File Access Error!"
End Sub
