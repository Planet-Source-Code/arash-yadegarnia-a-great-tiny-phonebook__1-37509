VERSION 5.00
Begin VB.Form FRM2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FRM2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label LBLabo3 
      Caption         =   "Arash Yadegarnia"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   450
      TabIndex        =   2
      Top             =   1575
      Width           =   3765
   End
   Begin VB.Label LBLabo2 
      AutoSize        =   -1  'True
      Caption         =   "Developed By :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   975
      Width           =   1725
   End
   Begin VB.Label LBLabo 
      AutoSize        =   -1  'True
      Caption         =   "Prisco Phonebook"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   600
      TabIndex        =   0
      Top             =   225
      Width           =   3510
   End
End
Attribute VB_Name = "FRM2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
