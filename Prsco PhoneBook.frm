VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cybersoft Prisco Phonebook"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   ClipControls    =   0   'False
   Icon            =   "Prsco PhoneBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cndbfiles 
      Left            =   600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FilterIndex     =   1
   End
   Begin VB.CommandButton CMD6 
      Caption         =   "Command1"
      Height          =   435
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Show The Timer"
      Top             =   2340
      Width           =   1155
   End
   Begin VB.CommandButton CMD5 
      Caption         =   "Command1"
      Height          =   435
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Load a Record"
      Top             =   1320
      Width           =   1155
   End
   Begin VB.CommandButton CMD4 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "About The Cybersoft"
      Top             =   4680
      Width           =   1155
   End
   Begin VB.CommandButton CMD3 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exit To Windows"
      Top             =   4680
      Width           =   1155
   End
   Begin VB.CommandButton CMD2 
      Caption         =   "Command1"
      Height          =   435
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Reset Record"
      Top             =   2340
      Width           =   1155
   End
   Begin VB.Frame FRA2 
      Caption         =   "Frame1"
      Height          =   2235
      Left            =   4080
      TabIndex        =   14
      Top             =   840
      Width           =   3195
      Begin VB.CommandButton CMD1 
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   435
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Save The Record"
         Top             =   480
         Width           =   1155
      End
   End
   Begin VB.TextBox TXT6 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   1
      ToolTipText     =   "Enter ID Number"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TXT4 
      Height          =   315
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   5
      ToolTipText     =   "Enter Area Code"
      Top             =   3240
      Width           =   675
   End
   Begin VB.Frame FRA1 
      Caption         =   "Frame1"
      Height          =   3435
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   3735
      Begin VB.TextBox TXT5 
         Height          =   315
         Left            =   1440
         MaxLength       =   21
         TabIndex        =   6
         ToolTipText     =   "Enter Phone Number"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox TXT2 
         Height          =   315
         Left            =   1440
         MaxLength       =   21
         TabIndex        =   3
         ToolTipText     =   "Enter Last Name"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox TXT3 
         Height          =   315
         Left            =   1440
         MaxLength       =   21
         TabIndex        =   4
         ToolTipText     =   "Enter State Or City"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox TXT1 
         Height          =   315
         Left            =   1440
         MaxLength       =   21
         TabIndex        =   2
         ToolTipText     =   "Enter First Name"
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label LBL7 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   480
      End
      Begin VB.Label LBL6 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2940
         Width           =   480
      End
      Begin VB.Label LBL5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label LBL3 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label LBL4 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label LBL2 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   480
      End
   End
   Begin VB.Label LBL9 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1320
      TabIndex        =   17
      Top             =   5580
      Width           =   480
   End
   Begin VB.Label LBL1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   75
   End
End
Attribute VB_Name = "FRM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub restore()
TXT6.Text = value.id
TXT1.Text = value.firstname
TXT2.Text = value.lastname
TXT3.Text = value.cityorstate
TXT4.Text = value.areacode
TXT5.Text = value.phonenumber
End Sub
Public Sub RAT()  'this subroutine sets all
TXT1.Text = ""     'textboxes empty.
TXT2.Text = ""
TXT3.Text = ""
TXT4.Text = ""
TXT5.Text = ""
TXT6.Text = ""
End Sub
Private Sub CMD1_Click()
Dim intfnum As Integer      'decleares all variables
Dim strfnam As String
On Error GoTo errhandler1
value.id = TXT6.Text
value.firstname = TXT1.Text
value.lastname = TXT2.Text
value.cityorstate = TXT3.Text
value.areacode = TXT4.Text
value.phonenumber = TXT5.Text
cndbfiles.Filter = "Cybersoft Prisco Phonebook (*.CPB) |*.CPB"
cndbfiles.FilterIndex = 1
cndbfiles.Flags = &H4 + &H2
cndbfiles.DialogTitle = "Save..."
cndbfiles.ShowSave
strfnam = cndbfiles.FileName
intfnum = FreeFile    'sets the targe file namber
Open strfnam For Binary As #intfnum 'opens the file
Put #intfnum, 1, value.id   'this code writes values to selescted file
Put #intfnum, 5, value.firstname
Put #intfnum, 29, value.lastname
Put #intfnum, 59, value.cityorstate
Put #intfnum, 74, value.areacode
Put #intfnum, 80, value.phonenumber
Close #intfnum              'closes the file
Call RAT               'all textboxes sets to empty
errhandler1:
    FRM1.Show
End Sub

Private Sub CMD2_Click()
Call RAT
End Sub

Private Sub CMD3_Click()
Dim stryon As String
stryon = MsgBox("Are you sure you want to quit ?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit to windows")
    If stryon = vbYes Then
        Unload FRM1
    Else
        FRM1.Show
    End If
End Sub
Private Sub CMD4_Click()
FRM2.Show       'opens the about window
End Sub

Private Sub CMD5_Click()
Dim intfnum As Integer
Dim strfn, strln, strsc, strac, strpn, strid As String
Dim strfnam As String
On Error GoTo errhandler1
cndbfiles.Filter = "Cybersoft Prisco Phonebook (*.CPB) |*.CPB"
cndbfiles.FilterIndex = 1
cndbfiles.Flags = &H4 + &H2
cndbfiles.DialogTitle = "Load..."
cndbfiles.ShowOpen
strfnam = cndbfiles.FileName
intfnum = FreeFile
Open strfnam For Binary As #intfnum
Get #intfnum, 1, value.id      'this code loads and restores values from the saved file
Get #intfnum, 5, value.firstname
Get #intfnum, 29, value.lastname
Get #intfnum, 59, value.cityorstate
Get #intfnum, 74, value.areacode
Get #intfnum, 80, value.phonenumber
Close #intfnum
Call restore
errhandler1:
    FRM1.Show
End Sub

Private Sub CMD6_Click()
FRM3.Show        'loads the timer window
End Sub

Private Sub Form_Load()
LBL1.Caption = "Welcome to Prisco® Phonebook"
FRA1.Caption = "Individual Information"
LBL2.Caption = "First Name : "
LBL3.Caption = "Last Name : "
LBL4.Caption = "State/City : "
LBL5.Caption = "Area Code : "         'sets captions and labels
LBL6.Caption = "Phone Number : "
LBL7.Caption = "ID Number : "
FRA2.Caption = "Control Option"
CMD1.Caption = "Save..."
CMD2.Caption = "Reset"
LBL9.Caption = "Cybersoft® Prisco Phonebook  -  All Rights Reserved For Developer"
CMD3.Caption = "Quit"
CMD4.Caption = "About.."
CMD5.Caption = "Load..."
CMD6.Caption = "Clock"
End Sub
Private Sub TXT5_Change()
If Not TXT5.Text = "" Then
    CMD1.Enabled = True
    Else: CMD1.Enabled = False
End If
End Sub
