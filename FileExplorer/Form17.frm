VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FFileInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Explorer"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   Icon            =   "Form17.frx":0000
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmGeneral 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   3120
      TabIndex        =   17
      Top             =   840
      Width           =   5295
      Begin VB.TextBox txtFilename 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "txtFilename"
         Top             =   360
         Width           =   3915
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "txtType"
         Top             =   1140
         Width           =   3435
      End
      Begin VB.TextBox txtLocation 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "txtLocation"
         Top             =   1440
         Width           =   3435
      End
      Begin VB.TextBox txtSize 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "txtSize"
         Top             =   1740
         Width           =   3435
      End
      Begin VB.TextBox txtCompSize 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "txtCompSize"
         Top             =   2040
         Width           =   3435
      End
      Begin VB.TextBox txtAccessed 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "txtAccessed"
         Top             =   3780
         Width           =   3435
      End
      Begin VB.TextBox txtModified 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "txtModified"
         Top             =   3480
         Width           =   3435
      End
      Begin VB.TextBox txtCreated 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "txtCreated"
         Top             =   3180
         Width           =   3435
      End
      Begin VB.TextBox txtDosName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "txtDosName"
         Top             =   2880
         Width           =   3435
      End
      Begin VB.TextBox txtDosPath 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "txtDosPath"
         Top             =   2580
         Width           =   3435
      End
      Begin VB.Frame frmAttributes 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1620
         TabIndex        =   19
         Top             =   4260
         Width           =   3255
         Begin VB.CheckBox chkAttr 
            Caption         =   "&Temporary"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   1620
            TabIndex        =   25
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox chkAttr 
            Caption         =   "&System"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   1620
            TabIndex        =   24
            Top             =   300
            Width           =   1335
         End
         Begin VB.CheckBox chkAttr 
            Caption         =   "Hi&dden"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   1620
            TabIndex        =   23
            Top             =   0
            Width           =   1335
         End
         Begin VB.CheckBox chkAttr 
            Caption         =   "Co&mpressed"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   22
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkAttr 
            Caption         =   "Ar&chive"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   21
            Top             =   300
            Width           =   1335
         End
         Begin VB.CheckBox chkAttr 
            Caption         =   "&Read-only"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   240
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   18
         Top             =   240
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   240
         X2              =   5100
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   240
         X2              =   5100
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   240
         X2              =   5100
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   240
         X2              =   5100
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   4
         X1              =   240
         X2              =   5100
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   5
         X1              =   240
         X2              =   5100
         Y1              =   4140
         Y2              =   4140
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   45
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label lblLocation 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   1740
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Compressed :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Label lblAccessed 
         AutoSize        =   -1  'True
         Caption         =   "Accessed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   3780
         Width           =   840
      End
      Begin VB.Label lblModified 
         AutoSize        =   -1  'True
         Caption         =   "Modified:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   3480
         Width           =   765
      End
      Begin VB.Label lblCreated 
         AutoSize        =   -1  'True
         Caption         =   "Created:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   3180
         Width           =   720
      End
      Begin VB.Label lblDosName 
         AutoSize        =   -1  'True
         Caption         =   "MS-DOS name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   2880
         Width           =   1245
      End
      Begin VB.Label lblDosPath 
         AutoSize        =   -1  'True
         Caption         =   "MS-DOS path:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label lblAttributes 
         AutoSize        =   -1  'True
         Caption         =   "Attributes:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   4260
         Width           =   915
      End
   End
   Begin VB.Frame frmVersion 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   5175
      Begin VB.TextBox txtFileVer 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "txtFileVer"
         Top             =   180
         Width           =   3615
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "txtDescription"
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtCopyright 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "txtCopyright"
         Top             =   1020
         Width           =   3615
      End
      Begin VB.Frame frmVerInfo 
         Caption         =   "Other version information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   240
         TabIndex        =   6
         Top             =   1500
         Width           =   4815
         Begin VB.TextBox txtVerInfo 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   2280
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Text            =   "Form17.frx":0CCA
            Top             =   600
            Width           =   2295
         End
         Begin VB.ListBox lstVerInfo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2220
            IntegralHeight  =   0   'False
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblVerItem 
            AutoSize        =   -1  'True
            Caption         =   "Item name:"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblVerValue 
            AutoSize        =   -1  'True
            Caption         =   "Value:"
            Height          =   195
            Left            =   2280
            TabIndex        =   9
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.Label lblFileVer 
         AutoSize        =   -1  'True
         Caption         =   "File version::"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         Caption         =   "Copyright:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1020
         Width           =   870
      End
   End
   Begin MSComctlLib.TabStrip tabInfo 
      Height          =   5775
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10186
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Version"
            Key             =   "Version"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Files"
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2595
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2595
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Hidden          =   -1  'True
         Left            =   120
         System          =   -1  'True
         TabIndex        =   1
         Top             =   2880
         Width           =   2595
      End
   End
End
Attribute VB_Name = "FFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' *********************************************************************
'  Copyright (C)1995-98 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  Warning: This computer program is protected by copyright law and
'  international treaties. Unauthorized reproduction or distribution
'  of this program, or any portion of it, may result in severe civil
'  and criminal penalties, and will be prosecuted to the maximum
'  extent possible under the law.
' *********************************************************************
Option Explicit

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

Private Const attrReadOnly = 0
Private Const attrArchive = 1
Private Const attrCompressed = 2
Private Const attrHidden = 3
Private Const attrSystem = 4
Private Const attrTemporary = 5

Private m_UserFile As String

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo err
   Dir1.Path = Drive1.Drive
   Exit Sub
err:
End Sub

Private Sub File1_Click()
   With File1
      If Right(.Path, 1) = "\" Then
         m_UserFile = .Path & .FileName
      Else
         m_UserFile = .Path & "\" & .FileName
      End If
   End With
   Call UpdateInfo(m_UserFile)
End Sub

Private Sub File1_PathChange()
   If File1.ListCount Then
      File1.ListIndex = 0
   Else
      m_UserFile = Dir1.Path
      Call UpdateInfo(m_UserFile)
   End If
End Sub

Private Sub Form_Load()
Dim Ret As Long
   Dim I As Long
   '
   ' Set initial dirspec
   '
   Drive1.Drive = Environ("windir")
   '
   ' Adjust 3d lines
   '
   For I = 1 To 5 Step 2
      Line1(I).Y1 = Line1(I - 1).Y1 + Screen.TwipsPerPixelY
      Line1(I).Y2 = Line1(I).Y1
   Next I
   '
   ' Make sure picture for icon is properly sized.
   '
   picIcon.Width = 32 * Screen.TwipsPerPixelX
   picIcon.Height = 32 * Screen.TwipsPerPixelY
   '
   ' Fill version info listbox
   '
   lstVerInfo.AddItem "Company Name"
   lstVerInfo.AddItem "Description"
   lstVerInfo.AddItem "Internal Name"
   lstVerInfo.AddItem "Language"
   lstVerInfo.AddItem "Legal Copyright"
   lstVerInfo.AddItem "Legal Trademarks"
   lstVerInfo.AddItem "Original Filename"
   lstVerInfo.AddItem "Product Name"
   lstVerInfo.AddItem "Product Version"
   '
   ' Position frames within tab
   '
   With tabInfo
      frmGeneral.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
      frmVersion.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
   End With
   frmGeneral.BackColor = Me.BackColor
   frmVersion.BackColor = Me.BackColor
   frmVersion.Visible = False
End Sub

Private Sub UpdateInfo(ByVal fil As String)
   Dim fi As CFileInfo
   Dim vi As CFileVersionInfo
   '
   ' Set current tab.
   '
   tabInfo.Tabs("General").Selected = True
   '
   ' Update all attribute information using intentionally
   ' mis-cased copy of m_UserFile
   '
   fil = UCase(fil)
   Set fi = New CFileInfo
   fi.FullPathName = fil
   '
   ' Fill controls with attributes.
   '
   txtFilename.Text = fi.DisplayName
   txtType.Text = fi.TypeName
   txtLocation = fi.FilePath
   txtSize.Text = fi.FormatFileSize(fi.FileSize)
   If fi.attrCompressed Then
      txtCompSize.Text = fi.FormatFileSize(fi.CompressedFileSize)
   Else
      txtCompSize.Text = "File is not compressed"
   End If
   txtDosPath.Text = fi.ShortPath
   txtDosName.Text = fi.ShortName
   txtCreated.Text = fi.FormatFileDate(fi.CreationTime)
   txtModified.Text = fi.FormatFileDate(fi.ModifyTime)
   txtAccessed.Text = fi.FormatFileDate(fi.LastAccessTime)
   chkAttr(attrReadOnly).Value = Abs(fi.attrReadOnly)
   chkAttr(attrArchive).Value = Abs(fi.attrArchive)
   chkAttr(attrCompressed).Value = Abs(fi.attrCompressed)
   chkAttr(attrHidden).Value = Abs(fi.attrHidden)
   chkAttr(attrSystem).Value = Abs(fi.attrSystem)
   chkAttr(attrTemporary).Value = Abs(fi.attrTemporary)
   '
   ' Display associated icon.
   '
   picIcon.Cls
   Call DrawIcon(picIcon.hdc, 0, 0, fi.hIcon)
   '
   ' Update version information
   '
   Set vi = New CFileVersionInfo
   vi.FullPathName = fi.FullPathName
   If vi.Available Then
      If tabInfo.Tabs.Count = 1 Then
         tabInfo.Tabs.Add 2, "Version", "Version"
      End If
      txtFileVer.Text = vi.FileVersion
      txtDescription.Text = vi.FileDescription
      txtCopyright.Text = vi.LegalCopyright
      lstVerInfo.ListIndex = 0
      lstVerInfo_Click
   Else
      If tabInfo.Tabs.Count > 1 Then
         tabInfo.Tabs.Remove 2
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstVerInfo_Click()
   Dim vi As New CFileVersionInfo
   vi.FullPathName = m_UserFile

   Select Case lstVerInfo.List(lstVerInfo.ListIndex)
      Case "Company Name"
         txtVerInfo.Text = vi.CompanyName
      Case "Description"
         txtVerInfo.Text = vi.FileDescription
      Case "Internal Name"
         txtVerInfo.Text = vi.InternalName
      Case "Language"
         txtVerInfo.Text = vi.Language
      Case "Legal Copyright"
         txtVerInfo.Text = vi.LegalCopyright
      Case "Legal Trademarks"
         txtVerInfo.Text = vi.LegalTrademarks
      Case "Original Filename"
         txtVerInfo.Text = vi.OriginalFilename
      Case "Product Name"
         txtVerInfo.Text = vi.ProductName
      Case "Product Version"
         txtVerInfo.Text = vi.ProductVersion
   End Select
End Sub

Private Sub tabInfo_Click()
   If tabInfo.Tabs("General").Selected Then
      frmVersion.Visible = False
      frmGeneral.Visible = True
   Else
      frmVersion.Visible = True
      frmGeneral.Visible = False
   End If
End Sub

