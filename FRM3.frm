VERSION 5.00
Begin VB.Form FRM3 
   Caption         =   "Smart Timer"
   ClientHeight    =   2100
   ClientLeft      =   12135
   ClientTop       =   3360
   ClientWidth     =   3705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TIM1 
      Left            =   2820
      Top             =   900
   End
   Begin VB.CommandButton CMD7 
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Return To Phonebook"
      Top             =   1380
      Width           =   1155
   End
   Begin VB.Label LBL8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1755
      TabIndex        =   0
      Top             =   300
      Width           =   165
   End
End
Attribute VB_Name = "FRM3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD7_Click()
Unload Me      'closes the about window
End Sub

Private Sub Form_Load()
TIM1.Interval = 1000     'sets the timer speed
CMD7.Caption = "<<  Back"   'labels and captions
LBL8.Caption = Time
End Sub
Private Sub TIM1_Timer()
LBL8.Caption = Time    'shows the timer in the label
End Sub





     'I spent 7 hours on writing this application
     'the results was acceptable for the first database
     '21/3/2002 ,1381/1/1     2:55 PM


'Arash Yadegarnia
