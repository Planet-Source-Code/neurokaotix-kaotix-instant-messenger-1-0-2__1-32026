VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Options 
   Caption         =   "Config"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Options.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
