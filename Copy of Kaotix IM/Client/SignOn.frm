VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SignOn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sign on to the Kaotix Network"
   ClientHeight    =   3735
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SignOn.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1200
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "JPEG (*.JPG)|*.jpg|Bitmap (*.BMP)|*.bmp|GIF (*.GIF)|*.gif|Windows Icons (*.ICO)|*.ico"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SignOn.frx":2CFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Accounts"
      TabPicture(0)   =   "SignOn.frx":3370
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "IList1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPassword"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LaVolpeButton2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "New ScreenNames"
      TabPicture(1)   =   "SignOn.frx":338C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin KIM.LaVolpeButton LaVolpeButton2 
         Height          =   375
         Left            =   645
         TabIndex        =   8
         Top             =   1260
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Click here to add an account!"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   99
         MICON           =   "SignOn.frx":33A8
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   2640
         Width           =   4265
      End
      Begin KIM.LaVolpeButton Command1 
         Height          =   540
         Left            =   3160
         TabIndex        =   3
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   953
         BTYPE           =   3
         TX              =   "Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   99
         MICON           =   "SignOn.frx":36C2
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin KIM.IList IList1 
         Height          =   1845
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Select an account"
         Top             =   480
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   3254
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New user? Create a ScreenName!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         MouseIcon       =   "SignOn.frx":39DC
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   3120
         Width           =   2910
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   315
         TabIndex        =   6
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Image Image3 
         Height          =   150
         Left            =   120
         Top             =   2445
         Width           =   150
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forgot Password?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2880
         MouseIcon       =   "SignOn.frx":3CE6
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   330
         TabIndex        =   4
         Top             =   2415
         Width           =   1575
      End
   End
   Begin VB.Image Green 
      Height          =   150
      Left            =   960
      Picture         =   "SignOn.frx":3FF0
      Top             =   4200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Red 
      Height          =   150
      Left            =   720
      Picture         =   "SignOn.frx":4069
      Top             =   4200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Menu mnuMyKIM 
      Caption         =   "&My KIM"
      Begin VB.Menu mnuMyKIMExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Accounts"
      Begin VB.Menu mnuAddAcount 
         Caption         =   "&Add New Account"
      End
      Begin VB.Menu mnuAccountSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuChangeIcon 
         Caption         =   "Change Acct. Icon"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About KIM"
      End
   End
End
Attribute VB_Name = "SignOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Set IList1.ImageList = ImageList1
IList1.ItemHeight = 40
IList1.SetPos 40, 4, 40, 20, 4, 4

IList1.AddItem "admin", "Double click to sign onto this account", , 1
IList1.AddItem "xpinkhex", "Double click to sign onto this account", , 1
IList1.AddItem "elementneon", "Double click to sign onto this account", , 1
'IList1.AddItem "admin", "*****", , 1

If IList1.Count <= 0 Then LaVolpeButton2.Visible = True:: txtPassword.Text = "": txtPassword.Enabled = False: txtPassword.BackColor = vbButtonFace: Image3.Picture = Red.Picture

Image3.Picture = Red.Picture
End Sub

Private Sub Command1_Click()
FinalClose = True
Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not FinalClose Then
Me.WindowState = 1
Cancel = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub IList1_DblClick()
    If IList1.Count <= 0 Then Exit Sub
    If txtPassword = "" Then MsgBox "No Password entered!", vbCritical: Exit Sub
    
    Dim Itm As CItem
    Set Itm = IList1.Item(IList1.Selected)
    YourSN = LCase(Itm.Caption)
    
    If Client.Winsock1.State <> sckClosed Then Client.Winsock1.Close
    Client.Winsock1.RemotePort = 1008
    'Client.Winsock1.RemoteHost = "216.77.225.246" 'put your IP here and comment out the one below
    Client.Winsock1.RemoteHost = "127.0.0.1"       'to allow people to connect to your IP
    Client.Winsock1.Connect
    
Do Until Client.Winsock1.State = sckConnected
    DoEvents: DoEvents: DoEvents: DoEvents
    If Client.Winsock1.State = sckError Then
        MsgBox "Could not connect to server! The server may be down or you may not be connected to the Internet. Check your connection and try again. If you still cannot connect wait until a later time when the server will be up."
        Exit Sub
    End If
Loop

    Client.Winsock1.SendData (".login" & " " & YourSN & " " & LCase(txtPassword.Text))
End Sub

Private Sub IList1_KeyDown(KeyCode As Integer, Shift As Integer)
If IList1.Count <= 0 Then Exit Sub
Select Case KeyCode
Case vbKeyDelete
mnuRemove_Click
End Select
End Sub



Private Sub Label5_Click()
Wizard.Show vbModal
End Sub

Private Sub LaVolpeButton2_Click()
mnuAddAcount_Click
End Sub

Private Sub mnuAddAcount_Click()
    Dim X2 As String
    
    X2 = InputBox("Enter ScreenName:", "Add new account")
    If X2 <> "" Then
        IList1.AddItem X2, "Double click to sign onto this account", , 1
        IList1.Redraw
        If IList1.Count >= 1 Then LaVolpeButton2.Visible = False: txtPassword.Enabled = True: txtPassword.BackColor = vbWhite: Image3.Picture = Red.Picture
    Else
    End If
End Sub

Private Sub mnuChangeIcon_Click()
    On Error GoTo BahError1
    Dim Itm As CItem
    Set Itm = IList1.Item(IList1.Selected)
    CD1.DialogTitle = "Select Account Icon (32x32, any other dimensions will not show correctly)"
    CD1.ShowOpen
    If CD1.FileName = "" Then Exit Sub
    ImageList1.ListImages.Add ImageList1.ListImages.Count + 1, , LoadPicture(CD1.FileName)
    Itm.Icon = ImageList1.ListImages.Count
    IList1.Redraw
    Exit Sub
    
BahError1:
    
    MsgBox "Error adding icon! One possible cause may be that the image is corrupted or not a valid format (e.g. a text file such as aaa.txt renamed to aaa.jpg).", vbCritical, "Could not load new account icon!"
    
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuMyKIMExit_Click()
FinalClose = True
Unload Me
End Sub

Private Sub txtPassword_Change()
If txtPassword.Text <> "" Then
Image3.Picture = Green.Picture
Else
Image3.Picture = Red.Picture
End If
End Sub

Private Sub IList1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IList1.Count <= 0 Then
    mnuRemove.Visible = False
    mnuRename.Visible = False
    mnuChangeIcon.Visible = False
    mnuAccountSep1.Visible = False
    If Button = 2 Then PopupMenu mnuEdit
    Else
    mnuRemove.Visible = True
    mnuRename.Visible = True
    mnuChangeIcon.Visible = True
    mnuAccountSep1.Visible = True
    If Button = 2 Then PopupMenu mnuEdit
    End If
End Sub

Private Sub IList1_OnSelect()
    'On Error Resume Next
    If x = 0 Then
        x = 1
        Exit Sub
    End If
End Sub

Private Sub mnuRemove_Click()
If IList1.Count <= 0 Then Exit Sub
Dim Itm As CItem
Set Itm = IList1.Item(IList1.Selected)
Dim Ans As Byte
Ans = MsgBox("Are you sure you want to remove the account '" & Itm.Caption & "'?", vbYesNo + vbQuestion, "Are you sure?")
If Ans = vbYes Then
IList1.Remove IList1.Selected
If IList1.Count <= 0 Then LaVolpeButton2.Visible = True: txtPassword.Text = "": txtPassword.Enabled = False: txtPassword.BackColor = vbButtonFace: Image3.Picture = Red.Picture
IList1.Redraw
Else
Exit Sub
End If
End Sub

Private Sub mnuRename_Click()
    Dim x As String
    Dim Itm As CItem
    Set Itm = IList1.Item(IList1.Selected)
    'On Error Resume Next

    x = InputBox("Enter new caption", , Itm.Caption)
    If x <> "" Then
        IList1.SetCaption IList1.Selected, x
    End If
    
End Sub


