VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Resolver"
   ClientHeight    =   2430
   ClientLeft      =   3390
   ClientTop       =   2655
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "IPresolve.frx":0000
   LinkTopic       =   "Form29"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.DNS DNS1 
      Left            =   240
      Top             =   1560
      _ExtentX        =   661
      _ExtentY        =   1085
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "localhost"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Solve"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address"
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Host Name"
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = DNS1.NameToAddress(Text2.Text)
Text2.Text = DNS1.AddressToName(Text1.Text)
End Sub

Private Sub Form_Load()
Text2.Text = GetIPHostName()
Text1.Text = GetIPAddress()
Command1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
Unload Me
End Sub

Private Sub Text1_Change()
Command1.Enabled = True
End Sub

Private Sub Text2_Change()
Command1.Enabled = True
End Sub
