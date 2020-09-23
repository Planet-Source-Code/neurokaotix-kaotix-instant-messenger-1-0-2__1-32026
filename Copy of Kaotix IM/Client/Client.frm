VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Client 
   ClientHeight    =   4575
   ClientLeft      =   13845
   ClientTop       =   3600
   ClientWidth     =   2295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Client.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleMode       =   0  'User
   ScaleWidth      =   2299.396
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Client.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Client.frx":0393
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Client.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Client.frx":04A5
            Key             =   "down"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Client.frx":0510
            Key             =   "right"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox MainFrame 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   2295
      TabIndex        =   1
      Top             =   480
      Width           =   2295
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   1080
         Top             =   2520
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tmrTimeout 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1560
         Top             =   2520
      End
      Begin TabDlg.SSTab SSTab 
         Height          =   4095
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   7223
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabMaxWidth     =   2117
         WordWrap        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Online"
         TabPicture(0)   =   "Client.frx":057E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "TreeView1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "runlog"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "List Setup"
         TabPicture(1)   =   "Client.frx":059A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TreeView2"
         Tab(1).Control(1)=   "cmdNewBuddy"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdDelBuddy"
         Tab(1).ControlCount=   3
         Begin VB.CommandButton cmdDelBuddy 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   495
            Left            =   -73680
            TabIndex        =   4
            Top             =   3480
            Width           =   855
         End
         Begin VB.CommandButton cmdNewBuddy 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -74880
            MaskColor       =   &H00FF00FF&
            Picture         =   "Client.frx":05B6
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   " Add a new buddy  "
            Top             =   3480
            UseMaskColor    =   -1  'True
            Width           =   735
         End
         Begin RichTextLib.RichTextBox runlog 
            Height          =   735
            Left            =   120
            TabIndex        =   5
            Top             =   3240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1296
            _Version        =   393217
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"Client.frx":08CA
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
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   2895
            Left            =   -74880
            TabIndex        =   6
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   5106
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LineStyle       =   1
            Style           =   6
            HotTracking     =   -1  'True
            ImageList       =   "ImageList1"
            Appearance      =   1
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
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   2655
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   4683
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   5
            HotTracking     =   -1  'True
            ImageList       =   "ImageList1"
            Appearance      =   1
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
      End
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Online"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      Top             =   90
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&My KIM"
      Begin VB.Menu mnuFileLogOut 
         Caption         =   "&Sign Off"
      End
      Begin VB.Menu mnuFileSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileStatus 
         Caption         =   "My &Status"
         Begin VB.Menu mnuStatusOnline 
            Caption         =   "&Online"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuStatusAway 
            Caption         =   "&Away"
         End
      End
      Begin VB.Menu mnuStatusSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuPeople 
      Caption         =   "&People"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About KIM"
      End
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "&Systray"
      Visible         =   0   'False
      Begin VB.Menu mnuSystraySignOff 
         Caption         =   "&Sign Off"
      End
      Begin VB.Menu mnuSystrayStatus 
         Caption         =   "My &Status"
         Begin VB.Menu mnuSystrayOnline 
            Caption         =   "&Online"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSystrayAway 
            Caption         =   "&Away"
         End
      End
      Begin VB.Menu mnuSystraySep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystrayExit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
   End
   Begin VB.Menu mnuBuddyList 
      Caption         =   "&BuddyList"
      Visible         =   0   'False
      Begin VB.Menu mnuRemBuddy 
         Caption         =   "&Remove Buddy"
      End
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strIncoming As String
Dim Start As Integer
Dim oldLabel As String
Dim Tic As NOTIFYICONDATA
Dim rc As Long

Private Sub cmdSignOn_Click()
    If Winsock1.State <> sckClosed Then Winsock1.Close
    Winsock1.RemotePort = 1008
    'Winsock1.RemoteHost = "216.77.225.246" 'put your IP here and comment out the one below
    Winsock1.RemoteHost = "127.0.0.1"       'to allow people to connect to your IP
    Winsock1.Connect
    
Do Until Winsock1.State = sckConnected
    DoEvents: DoEvents: DoEvents: DoEvents
    If Winsock1.State = sckError Then
        MsgBox "Problem connecting!"
        Exit Sub
    End If
Loop
    Winsock1.SendData (".login" & " " & LCase(cmbUsername.Text) & " " & LCase(txtPassword.Text))
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, Tic
End Sub

Private Sub lblStatus_Click()
    PopupMenu mnuFileStatus
End Sub

Private Sub DoSystrayIcon()
    Tic.cbSize = Len(Tic)
    Tic.hWnd = Me.hWnd
    Tic.uID = vbNull
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Me.Icon
    Tic.sTip = "Kaotix Instant Messenger (" & YourSN & ")" & vbNullChar
    rc = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        Dim msg As Long
        Dim sFilter As String
        msg = x / Screen.TwipsPerPixelX
        Select Case msg
            Case WM_RBUTTONUP
                PopupMenu mnuSystray
            Case WM_LBUTTONDBLCLK
                Me.Show
                'Me.WindowState = vbNormal
        End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Winsock1.State <> sckClosed Then Winsock1.Close
    End
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.Height < 4000 Then
        Me.Height = 4000
        Exit Sub
    End If
    
    MainFrame.Width = Me.ScaleWidth
    MainFrame.Height = Me.ScaleHeight
    
    SSTab.Width = MainFrame.Width '- 65
    SSTab.Height = MainFrame.Height - 500
    
    TreeView1.Width = SSTab.Width - 260
    TreeView1.Height = SSTab.Height - 1350
    TreeView2.Width = TreeView1.Width
    TreeView2.Height = TreeView1.Height
    
    runlog.Width = TreeView1.Width
    runlog.Top = TreeView1.Top + TreeView1.Height + 25
    
    cmdNewBuddy.Top = SSTab.Height - 625
    cmdDelBuddy.Top = SSTab.Height - 525
    
    Shape1.Width = Me.ScaleWidth
    
    lblStatus.Left = Me.ScaleWidth - lblStatus.Width - 120
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuRestore_Click()
Me.Show
'Me.WindowState = vbNormal
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Bold = False
    Node.Image = "right"
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
    Node.Bold = True
    Node.Image = "down"
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Expanded = False Then
    Node.Expanded = True
Else
    Node.Expanded = False
End If
If Not Node.Parent Is Nothing Then
    If Node.Image <> 2 Then
            Dim test As Integer
            FormIsLoaded (TreeView1.SelectedItem)
    End If
End If
End Sub

Private Function FormIsLoaded(frm As String)
Dim FormNbr As Integer
Dim flag As Integer
flag = 0
For FormNbr = 0 To Forms.Count - 1
  If LCase(Word(Forms(FormNbr).Caption, 1)) = LCase(frm) Then
        Forms(FormNbr).SetFocus
        Exit Function
  Else
    flag = 1
  End If
Next FormNbr

If flag = 1 Then
    Dim NewIMessage As New IMessage
    NewIMessage.Show ownerform:=Me
    NewIMessage.Caption = frm & " - Instant Message"
End If
End Function

Private Function GetFormNumber(frm As String) As Integer
Dim FormNbr As Integer
For FormNbr = 0 To Forms.Count - 1
  If LCase(Word(Forms(FormNbr).Caption, 1)) = LCase(frm) Then
        GetFormNumber = FormNbr
        Exit Function
 End If
Next FormNbr
End Function

Private Sub TreeView2_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Bold = False
    cmdDelBuddy.Enabled = False
End Sub

Private Sub TreeView2_Expand(ByVal Node As MSComctlLib.Node)
    Node.Bold = True
End Sub

Private Sub TreeView2_BeforeLabelEdit(Cancel As Integer)
    oldLabel = TreeView2.SelectedItem.Text
End Sub

Private Sub TreeView2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then PopupMenu mnuBuddyList
End Sub


Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Text <> "Buddies" Then
        cmdDelBuddy.Enabled = True
    Else
        cmdDelBuddy.Enabled = False
    End If
End Sub

Private Sub TreeView2_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim tn As Node
If oldLabel <> "Buddies" Then
    If UCase(NewString) <> UCase(oldLabel) Then
        If Correct_Screenname(NewString) = True Then
            If check_for_duplicate(NewString) = True Then
                TreeView2.SelectedItem.Key = NewString
                For Each tn In TreeView1.Nodes
                    If UCase(tn.Key) = UCase(oldLabel) Then
                        tn.Key = NewString
                        tn.Text = NewString
                        Winsock1.SendData ".updateBuddy " & YourSN & " " & oldLabel & " " & NewString
                        WaitFor (".statusUpdate")
                        Exit For
                    End If
                Next
            Else
                MsgBox "A buddy with the user name " & UCase(NewString) & " already exists.", vbOKOnly + vbCritical
                Cancel = 1
            End If
            Else
                Cancel = 1
            End If
        Else
            Cancel = 1
        End If
    Else
        Cancel = 1
End If
End Sub

Private Function check_for_duplicate(user As String) As Boolean
Dim tn As Node
Dim flag As Integer
    For Each tn In TreeView1.Nodes
        If UCase(tn.Key) = UCase(user) Then
            flag = 1
            check_for_duplicate = False
            Exit For
        Else
            flag = 0
        End If
    Next
If flag = 0 Then
    check_for_duplicate = True
End If
End Function

Private Sub cmdDelBuddy_Click()
Dim reply As String
Dim tn As Node
    reply = MsgBox("Are you sure you want to delete the following buddy from your list?" & vbCrLf & TreeView2.SelectedItem.Key, vbYesNo + vbCritical)
    If reply = vbYes Then
       Winsock1.SendData ".delBuddy " & YourSN & " " & TreeView2.SelectedItem.Key
       WaitFor (".statusUpdate")
        TreeView1.Nodes.Remove TreeView2.SelectedItem.Key
        TreeView2.Nodes.Remove TreeView2.SelectedItem.Key
       cmdDelBuddy.Enabled = False
       Call Online_Offline_Text
    End If
End Sub

Private Sub cmdNewBuddy_Click()
Dim newbuddy As String
    cmdNewBuddy.Enabled = False
    newbuddy = InputBox("Enter user name:", cmdNewBuddy.ToolTipText)
    If StrPtr(newbuddy) = 0 Then
        cmdNewBuddy.Enabled = True
    Else
        If Correct_Screenname(newbuddy) = True Then
            If check_for_duplicate(newbuddy) = True Then
                TreeView1.Nodes.Add "Offline", tvwChild, newbuddy, newbuddy
                TreeView2.Nodes.Add "Buddies", tvwChild, newbuddy, newbuddy
                Winsock1.SendData ".newBuddy " & YourSN & " " & newbuddy
                WaitFor (".statusUpdate")
                cmdNewBuddy.Enabled = True
            Else
               MsgBox "A buddy with the user name " & UCase(newbuddy) & " already exists.", vbOKOnly + vbCritical
               Call cmdNewBuddy_Click
            End If
        Else
            Call cmdNewBuddy_Click
        End If
    End If
End Sub

Private Function Correct_Screenname(screenname As String) As Boolean
Dim i As Integer
Dim flag As Integer
If LCase(screenname) <> YourSN And Len(screenname) >= 5 And Len(screenname) <= 15 And Not IsNumeric(Left(screenname, 1)) Then
For i = 1 To Len(screenname)
    If (Asc(Mid(screenname, i, 1)) >= vbKey1 And Asc(Mid(screenname, i, 1)) <= vbKey9) Then
        flag = 0
    ElseIf (Asc(UCase(Mid(screenname, i, 1))) >= vbKeyA And Asc(UCase(Mid(screenname, i, 1))) <= vbKeyZ) Then
        flag = 0
    Else
        flag = 1
        MsgBox "A screen name in your list is too short or contains invalid" & vbCrLf & "characters.", vbOKOnly + vbCritical
        Correct_Screenname = False
        Exit For
    End If
Next i
Else
    flag = 1
    MsgBox "A screen name in your list is too short or contains invalid" & vbCrLf & "characters.", vbOKOnly + vbCritical
    Correct_Screenname = False
End If
If flag = 0 Then
    Correct_Screenname = True
End If
End Function

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim i As Long
    Winsock1.GetData strIncoming
   
    If strIncoming = ".badlogin" Then
        MsgBox "The screen name or password you entered is not valid. ", vbOKOnly + vbCritical: SignOn.Image3.Picture = SignOn.Red.Picture
        If Winsock1.State <> sckClosed Then
            Winsock1.Close
        End If
    ElseIf strIncoming = ".goodlogin" Then
        Call good_login
        
    ElseIf Word(strIncoming, 1) = ".showonline" And Word(strIncoming, 2) <> "0" Then
        Call Show_Online_buddies(strIncoming)
        
    ElseIf Word(strIncoming, 1) = ".statusUpdate" And Word(strIncoming, 2) <> "0" Then
        Call status_update(Word(strIncoming, 3), Word(strIncoming, 4))
        
    ElseIf Word(strIncoming, 1) = ".msg" Then
        Call get_message(Word(strIncoming, 2), strIncoming)
        
  ' ElseIf Word(strIncoming, 1) = ".define" Then
       ' Call get_definition(Word(strIncoming, 2), Word(strIncoming, 3), strIncoming)
        
   ' ElseIf Word(strIncoming, 1) = ".spell" Then
        '    Call get_spelling(Word(strIncoming, 2), strIncoming)

    End If
End Sub

'Private Function get_spelling(buddy As String, msg As String)
'Dim formNum As Integer
'Dim definition As String
'formNum = GetFormNumber(buddy)
'If formNum <> 0 Then
'    definition = SplitString(msg, "..//..")
'    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
'    Forms(formNum).showmsg.SelBold = True
'    Forms(formNum).showmsg.SelText = "spellcheck: "
    
'    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
'    Forms(formNum).showmsg.SelBold = False
'    Forms(formNum).showmsg.SelText = definition & vbCrLf
'End If
'End Function

'Private Function get_definition(buddy As String, Word As String, msg As String)
'Dim formNum As Integer
'Dim definition As String
'formNum = GetFormNumber(buddy)
'If formNum <> 0 Then
'    definition = SplitString(msg, "..//..")
'    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
'    Forms(formNum).showmsg.SelBold = True
'    Forms(formNum).showmsg.SelText = Word & ": "
    
'    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
'    Forms(formNum).showmsg.SelBold = False
'    Forms(formNum).showmsg.SelText = definition & vbCrLf
'End If
'End Function

Private Function get_message(mfrom As String, msg As String)
Dim i As Long
Dim formNum As Integer
Dim sendMsg As String
    sndPlaySound App.Path + "\sounds\imrcv.wav", 1
    FormIsLoaded (mfrom)
    formNum = GetFormNumber(mfrom)
    sendMsg = SplitString(msg, "..//..")
    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
    Forms(formNum).showmsg.SelBold = True
    Forms(formNum).showmsg.SelColor = vbBlue
    Forms(formNum).showmsg.SelText = mfrom & ": "
    
    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
    Forms(formNum).showmsg.SelBold = False
    Forms(formNum).showmsg.SelColor = vbBlack
    Forms(formNum).showmsg.SelText = sendMsg & vbCrLf
End Function

Private Sub good_login()

        Me.Caption = YourSN
        Me.Caption = Me.Caption + " - Kaotix Instant Messenger"
        
        SignOn.Visible = False
        Client.Visible = True
        
        DoSystrayIcon
        
        strIncoming = ""
            TreeView1.Nodes.Add , , "Online", "Buddies", "down"
            TreeView1.Nodes.Add , , "Offline", "Offline", "down"
            TreeView2.Nodes.Add , , "Buddies", "Buddies"
            
            TreeView1.Nodes.Item(1).Expanded = True
            TreeView1.Nodes.Item(2).Expanded = True
            TreeView2.Nodes.Item(1).Expanded = True
            
            TreeView1.Nodes.Item(1).Bold = True
            TreeView2.Nodes.Item(1).Bold = True
            
            TreeView1.Nodes.Item(2).ForeColor = vbButtonShadow
            
                        
            Winsock1.SendData ".updateStatus" & " " & "1" & " " & YourSN
            WaitFor (".statusUpdate")
                       
            Winsock1.SendData ".getonlinebuddies" & " " & YourSN
            WaitFor (".showonline")
End Sub

Private Function Show_Online_buddies(buddies As String)
    Dim i As Long, oncount As Integer
    Dim status As Integer
    Dim n As String
    oncount = 0
    For i = 3 To Words(strIncoming)
        status = Right(Word(buddies, i), 1)
        If status = 2 Then
            n = "Offline"
        Else
            n = "Online"
        End If
        TreeView1.Nodes.Add n, tvwChild, Left(Word(buddies, i), Len(Word(buddies, i)) - 1), Left(Word(buddies, i), Len(Word(buddies, i)) - 1), status, status
        TreeView2.Nodes.Add "Buddies", tvwChild, Left(Word(buddies, i), Len(Word(buddies, i)) - 1), Left(Word(buddies, i), Len(Word(buddies, i)) - 1)
    Next i
    Call Online_Offline_Text
End Function

Private Function status_update(buddy As String, status As Integer)
    Dim tn As Node
    Dim n As String
    Dim frmNum As Integer

        If status = 2 Then
            n = "Offline"
        Else
            n = "Online"
        End If
    For Each tn In TreeView1.Nodes
        If LCase(tn.Key) = LCase(buddy) Then
            'TreeView1.Nodes.Remove tn.Key
            'TreeView1.Nodes.Add n, tvwChild, buddy, buddy, status, status
            If Word(TreeView1.Nodes(tn.Key).Parent, 1) <> "Buddies" Then
                If status = 1 Then
                    sndPlaySound App.Path + "\sounds\dooropen.wav", 1
                    runlog.SelStart = Len(runlog.Text)
                    runlog.SelColor = vbBlue
                    runlog.SelText = buddy & " has signed on (" & Time & ")" & vbCrLf
                End If
                If status <> 2 Then
                    frmNum = GetFormNumber(LCase(buddy))
                    If frmNum <> 0 Then
                        Forms(frmNum).cmdSend.Enabled = True
                        Forms(frmNum).typemsg.Enabled = True
                        Forms(frmNum).tbFonts.Enabled = True
                        Forms(frmNum).showmsg.SelStart = Len(Forms(frmNum).showmsg.Text)
                        Forms(frmNum).showmsg.SelColor = vbBlue
                        Forms(frmNum).showmsg.SelText = buddy & " has signed on (" & Time & ")." & vbCrLf
                    End If
                End If
            End If
            If Word(TreeView1.Nodes(tn.Key).Parent, 1) <> "Offline" Then
                If status = 2 Then
                    sndPlaySound App.Path + "\sounds\doorslam.wav", 1
                    runlog.SelStart = Len(runlog.Text)
                    runlog.SelColor = vbRed
                    runlog.SelText = buddy & " has signed off (" & Time & ")" & vbCrLf
                    frmNum = GetFormNumber(LCase(buddy))
                    If frmNum <> 0 Then
                        Forms(frmNum).cmdSend.Enabled = False
                        Forms(frmNum).typemsg.Enabled = False
                        Forms(frmNum).tbFonts.Enabled = False
                        Forms(frmNum).showmsg.SelStart = Len(Forms(frmNum).showmsg.Text)
                        Forms(frmNum).showmsg.SelColor = vbRed
                        Forms(frmNum).showmsg.SelText = buddy & " has signed off (" & Time & ")." & vbCrLf
                    End If
                End If
            End If
            TreeView1.Nodes.Remove tn.Key
            TreeView1.Nodes.Add n, tvwChild, buddy, buddy, status, status
            Call Online_Offline_Text
            Exit For
        End If
    Next
End Function

Private Function Online_Offline_Text()
Dim tn As Node
Dim oncount
Dim offcount
oncount = 0
offcount = 0
TreeView1.Nodes.Item(1).Selected = True
    oncount = TreeView1.SelectedItem.Children
TreeView1.Nodes.Item(2).Selected = True
    offcount = TreeView1.SelectedItem.Children

    TreeView1.Nodes.Item(1).Text = "Buddies (" & oncount & "/" & oncount + offcount & ")"
    TreeView1.Nodes.Item(2).Text = "Offline (" & offcount & "/" & oncount + offcount & ")"
End Function

Private Sub mnuFileLogOut_Click()
    If Winsock1.State <> sckClosed Then Winsock1.Close
    Me.Width = 2700
    Me.Height = 6000
    Client.Visible = False
    SignOn.Visible = True
    TreeView1.Nodes.Clear
    TreeView2.Nodes.Clear
    SignOn.txtPassword.Text = ""
    mnuStatusAway.Checked = False
    mnuStatusOnline.Checked = True
Dim i As Integer
    For i = 0 To Forms.Count - 1
    If Forms.Count <> 1 Then
        Unload Forms(1)
    End If
    Next i
End Sub

Private Sub mnuFileClose_Click()
Dim Ans2 As Byte
Ans2 = MsgBox("Are you sure you would like to end this session?", vbYesNo + vbQuestion, "Are you sure?")
If Ans2 = vbYes Then
FinalClose = True
Shell_NotifyIcon NIM_DELETE, Tic
Unload Me
Else
Exit Sub
End If
End Sub

Private Sub mnuStatusOnline_Click()
    If mnuStatusOnline.Checked = False Then
        mnuStatusOnline.Checked = True
        mnuStatusAway.Checked = False
        lblStatus.Caption = "Online"
        Winsock1.SendData ".updateStatus" & " " & "1" & " " & YourSN
        WaitFor (".statusUpdate")
    End If
End Sub

Private Sub mnuStatusAway_Click()
    If mnuStatusAway.Checked = False Then
        mnuStatusAway.Checked = True
        mnuStatusOnline.Checked = False
        lblStatus.Caption = "Away"
        Winsock1.SendData ".updateStatus" & " " & "3" & " " & YourSN
        WaitFor (".statusUpdate")
    End If
End Sub

Sub WaitFor(ResponseCode As String)
    Start = 0
    tmrTimeout.Enabled = True
    While Len(strIncoming) = 0
        DoEvents
        If Start > 20 Then
            MsgBox "Service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
            Call mnuFileLogOut_Click
        End If
    Wend
    Start = 0
    While Word(strIncoming, 1) <> ResponseCode
        DoEvents
        If Start > 20 Then
           MsgBox "Service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + strIncoming, 64, MsgTitle
           Exit Sub
           Call mnuFileLogOut_Click
        End If
    Wend
    strIncoming = ""
    tmrTimeout.Enabled = False
End Sub

Private Sub tmrTimeout_Timer()
    Start = Start + 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not FinalClose Then
'Me.WindowState = 1
Me.Hide
Cancel = 1
End If
End Sub
