VERSION 5.00
Begin VB.Form Wizard 
   Caption         =   "New KIM ScreenName Wizard"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Wizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text4 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Next"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Finish"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      MaxLength       =   16
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3105
      ScaleWidth      =   1545
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      Height          =   195
      Left            =   1680
      TabIndex        =   11
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter desired password (Between 6-16 chars)"
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   1320
      Width           =   3930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your e-mail address (Required)"
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   720
      Width           =   3210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the screenname you wish to sign-up for (Between 3-16 chars)"
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   5835
   End
End
Attribute VB_Name = "Wizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
