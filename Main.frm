VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secure Me (+) Update! 3.1 FINAL!"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   840
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   975
      Left            =   0
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   1720
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Main.frx":0442
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "21,23,25,1080,8080"
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   3120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox newlog 
      Height          =   1905
      Left            =   0
      TabIndex        =   14
      Top             =   270
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   3360
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Main.frx":050B
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   7440
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Options"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save log to disk"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "View Soft Log"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View Hard Log"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label13 
      Height          =   495
      Left            =   2520
      TabIndex        =   23
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label12 
      Height          =   495
      Left            =   2520
      TabIndex        =   22
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label11 
      Height          =   495
      Left            =   2520
      TabIndex        =   21
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   2520
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label9 
      Height          =   495
      Left            =   2520
      TabIndex        =   19
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Height          =   495
      Left            =   2520
      TabIndex        =   18
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   2520
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2415
      Width           =   6120
   End
   Begin VB.Label Label2 
      Caption         =   "Port to listen for connection:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   15
      Width           =   2055
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "2"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'oh I am meth0s
'eat me!
'This program was Anti Hexxed by my program. download it here.
'http://www.planet-source-code.com/xq/ASP/txtCodeId.24314/lngWId.1/qx/vb/scripts/ShowCode.htm

Private Sub Command1_Click()
On Error GoTo er
'When there is a error. goto er at the end of the sub to find out what happens
'this is our start button when its clicked
Dim helloa As String
Dim hellob As String
Dim helloc As Integer
Dim hellod As String
Dim hello1 As String
'we are going to dim everything

hellod = ""
helloa = ""
hellob = ""
helloc = 0
hello1 = ""
List1.Clear
'we are going to clear these

Timer2.Enabled = False
'We disable this for error handling.
'IE if you where to push start then stop then start you would get a socket error.
'Note: IE means Example

helloa = Len(Text3.Text)
'saying helloa is however many characters are in text3
'Note: the len command can be used to find out how many characters are in a string.

For i = 1 To helloa
'For every character in helloa.
'i is nothing of importance. you can use Z instead of i you can use s instead of i
'it doesnt matter.

helloc = helloc + 1
'You will understand this later on in the code.

hellob = Mid(Text3.Text, i, 1)
'hellob is going to be the character of i which is the current character.
'Mid Returns specified number of characters from a string

If hellob = Chr(44) Then
'if hellob is , we are going to parse it.
'we are using the char command. so Chr(44) is ,
'So if the current character is , then we are going to do.

hellob = ""
'we are clearing away the , so that we only have the number.

hellod = hellod + hellob
'hellod is going to be the same as it use to be just add hellob to it.

If hellod = "" Then
'if hellod is nothing. we dont want to add it.
'If the user just put a bunch of ,,,,,,, and we didnt have this.
'then it would add just a bunch of nothing's to the list box
'so if hellod is nothing. then we do nothing =)

Else
'if it isnt nothing. then its going to be a successfull port.
'so we are going to add it to the list box of ports we want to open.

List1.AddItem hellod
'lets go ahead and add it to our listbox

End If
'Ending our if

hellod = ""
'clearing hellod so we start over on the next port.

Else
'if its not , then add it to hellod

hellod = hellod + hellob
'The number what it may be is going to be the same just add the current number
'on to it. so if it use to be 1 and the hellob is 2 then the port is going to be 12

If helloc = helloa Then List1.AddItem hellod
'remember up there. helloc = helloc + 1 well this is where we use it.
'if we are on the last character. we will see if its the last port.
'so we can add it without having to put the , for the parsing.

End If
'ending our very first if.

Next i
'go to the next character.

Call xListKillDupes(List1) 'calls sub from module
'This way we can kill everything that there is more then one of in the list box.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'now we will count all the winsock controls we need to load.'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

hello1 = List1.ListCount
'Hello1 is a string. and its going to be how ever many items are in our listbox

For c = 0 To List1.ListCount - 1
Load Winsock(c + 1) 'loading a new winsock control in
List1.ListIndex = c 'going to the current port in the listbox
Winsock(c).LocalPort = List1.Text 'assigning this new controls port.
Winsock(c).Listen 'making the new control listen on its new port.
Winsock(c).Tag = List1.Text 'so when they connect we know what port they connect on.
Next c ' Going to the next c
Command1.Enabled = False 'disabling start button so nothing can mess up!
Command2.Enabled = True 'Enabling the stop button so you can stop it!

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'All rich textbox stuff. simple enough to figure out works just like a textbox'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Everytime we want to change the color of the text we need to set the new color

newlog.Text = ""
newlog.SelColor = &H80000002
newlog.SelText = Chr(33) & Chr(33) & Chr(33)
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & Chr(32) & Chr(83) & Chr(116) & Chr(97) & Chr(114) & Chr(116) & Chr(101) & Chr(100) & Chr(32) & Chr(111) & Chr(110) & Chr(32)
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, Chr(71) & Chr(101) & Chr(110) & Chr(101) & Chr(114) & Chr(97) & Chr(108) & Chr(32) & Chr(68) & Chr(97) & Chr(116) & Chr(101))
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(32) & Chr(33) & Chr(33) & Chr(33) & vbCrLf
newlog.SelColor = &H80&
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(33) & Chr(33) & Chr(33) & Chr(32)
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & Chr(76) & Chr(105) & Chr(115) & Chr(116) & Chr(101) & Chr(110) & Chr(105) & Chr(110) & Chr(103) & Chr(32) & Chr(102) & Chr(111) & Chr(114) & Chr(32) & Chr(99) & Chr(111) & Chr(110) & Chr(110) & Chr(101) & Chr(99) & Chr(116) & Chr(105) & Chr(111) & Chr(110) & Chr(115) & Chr(32) & Chr(111) & Chr(110) & Chr(32) & Chr(84) & Chr(67) & Chr(80) & Chr(32) & Chr(112) & Chr(111) & Chr(114) & Chr(116) & Chr(40) & Chr(115) & Chr(41) & Chr(32)
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Text3.Text
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(32) & Chr(33) & Chr(33) & Chr(33) & vbCrLf
newlog.SelColor = &H80&

'all the Chr(xx) stuff was done using my Chr command program get it at.
'http://www.planet-source-code.com/xq/ASP/txtCodeId.24314/lngWId.1/qx/vb/scripts/ShowCode.htm

er:
If Err.Number = 10055 Then
MsgBox Chr(89) & Chr(111) & Chr(117) & Chr(32) & Chr(104) & Chr(97) & Chr(118) & Chr(101) & Chr(32) & Chr(114) & Chr(97) & Chr(110) & Chr(32) & Chr(111) & Chr(117) & Chr(116) & Chr(32) & Chr(111) & Chr(102) & Chr(32) & Chr(98) & Chr(117) & Chr(102) & Chr(102) & Chr(101) & Chr(114) & Chr(32) & Chr(115) & Chr(112) & Chr(97) & Chr(99) & Chr(101) & Chr(46) & vbCrLf & Chr(83) & Chr(111) & Chr(109) & Chr(101) & Chr(32) & Chr(111) & Chr(102) & Chr(32) & Chr(116) & Chr(104) & Chr(101) & Chr(32) & Chr(112) & Chr(111) & Chr(114) & Chr(116) & Chr(115) & Chr(32) & Chr(121) & Chr(111) & Chr(117) & Chr(32) & Chr(104) & Chr(97) & Chr(118) & Chr(101) & Chr(32) & Chr(116) & Chr(114) & Chr(105) & Chr(101) & Chr(100) & Chr(32) & Chr(116) & Chr(111) & Chr(32) & Chr(111) & Chr(112) & Chr(101) & Chr(110) & vbCrLf & Chr(87) & Chr(104) & Chr(101) & Chr(114) & Chr(101) & Chr(32) & Chr(110) & Chr(111) & Chr(116) & Chr(32) & Chr(111) & Chr(112) & Chr(101) & Chr(110) & Chr(101) & Chr(100) & Chr(32) & Chr(117) & Chr(112) & Chr(46) _
, vbCritical, Chr(87) & Chr(97) & Chr(114) & Chr(110) & Chr(105) & Chr(110) & Chr(103) & Chr(32) & Chr(45) & Chr(32) & Chr(87) & Chr(105) & Chr(110) & Chr(57) & Chr(120)

Command1.Enabled = False 'disabling start button so nothing can mess up!
Command2.Enabled = True 'Enabling the stop button so you can stop it!
newlog.SelColor = &H80000002
newlog.SelText = Chr(33) & Chr(33) & Chr(33)
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & Chr(32) & Chr(83) & Chr(116) & Chr(97) & Chr(114) & Chr(116) & Chr(101) & Chr(100) & Chr(32) & Chr(111) & Chr(110) & Chr(32)
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, Chr(71) & Chr(101) & Chr(110) & Chr(101) & Chr(114) & Chr(97) & Chr(108) & Chr(32) & Chr(68) & Chr(97) & Chr(116) & Chr(101))
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(32) & Chr(33) & Chr(33) & Chr(33) & vbCrLf
newlog.SelColor = &H80&
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(33) & Chr(33) & Chr(33) & Chr(32)
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & Chr(76) & Chr(105) & Chr(115) & Chr(116) & Chr(101) & Chr(110) & Chr(105) & Chr(110) & Chr(103) & Chr(32) & Chr(102) & Chr(111) & Chr(114) & Chr(32) & Chr(99) & Chr(111) & Chr(110) & Chr(110) & Chr(101) & Chr(99) & Chr(116) & Chr(105) & Chr(111) & Chr(110) & Chr(115) & Chr(32) & Chr(111) & Chr(110) & Chr(32) & Chr(116) & Chr(99) & Chr(112) & Chr(32) & Chr(112) & Chr(111) & Chr(114) & Chr(116) & Chr(40) & Chr(115) & Chr(41) & Chr(32)
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Text3.Text
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(32) & Chr(33) & Chr(33) & Chr(33) & vbCrLf
newlog.SelColor = &H80&
ElseIf Err.Number = 0 Then

Command1.Enabled = False 'disabling start button so nothing can mess up!
Command2.Enabled = True 'Enabling the stop button so you can stop it!
Else
MsgBox Err.Description, vbCritical, Chr(69) & Chr(114) & Chr(114) & Chr(111) & Chr(114)
End
'ending this so they can get a fresh start.
End If
newlog.SetFocus
End Sub

Private Sub Command2_Click()
List1.Clear
Timer2.Enabled = False
Timer3.Enabled = False
' this is our stop button when its clicked

Command1.Enabled = True 'Enabling the start button so it can be used again!
Command2.Enabled = False 'Disabling the stop button so nothing can mess up!
For i = Winsock.LBound + 1 To Winsock.UBound
Winsock(i).Close
Unload Winsock(i)
Next i
Winsock(0).Close
'A better winsock unloader. It was having trouble previously.
newlog.SelColor = &H80000002
newlog.SelText = Chr(33) & Chr(33) & Chr(33)
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & newlog.SelText & Chr(32) & Chr(83) & Chr(116) & Chr(111) & Chr(112) & Chr(112) & Chr(101) & Chr(100) & Chr(32) & Chr(111) & Chr(110) & Chr(32)
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, Chr(71) & Chr(101) & Chr(110) & Chr(101) & Chr(114) & Chr(97) & Chr(108) & Chr(32) & Chr(68) & Chr(97) & Chr(116) & Chr(101))
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(32) & Chr(33) & Chr(33) & Chr(33) & vbCrLf
newlog.SelColor = &H80&
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(33) & Chr(33) & Chr(33) & Chr(32)
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & Chr(67) & Chr(108) & Chr(111) & Chr(115) & Chr(105) & Chr(110) & Chr(103) & Chr(32) & Chr(84) & Chr(67) & Chr(80) & Chr(32) & Chr(112) & Chr(111) & Chr(114) & Chr(116) & Chr(40) & Chr(115) & Chr(41) & Chr(32)
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Text3.Text
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(32) & Chr(33) & Chr(33) & Chr(33) & vbCrLf
newlog.SelColor = &H80&
newlog.SetFocus
End Sub

Private Sub Command3_Click()
If Form1.Height = 4075 Then
Command3.Enabled = False
Command4.Enabled = True
Else
Form1.Height = Form1.Height + 1000
Clipboard.SetText Form1.Height
Text1.Visible = True
Label4.Caption = 1
Command3.Enabled = False
Command4.Enabled = True
newlog.SelColor = &H80000002
newlog.SelText = Chr(33) & Chr(33) & Chr(33) & Chr(32)
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, Chr(71) & Chr(101) & Chr(110) & Chr(101) & Chr(114) & Chr(97) & Chr(108) & Chr(32) & Chr(68) & Chr(97) & Chr(116) & Chr(101))
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & Chr(32) & Chr(45) & Chr(32) & Chr(72) & Chr(97) & Chr(114) & Chr(100) & Chr(32) & Chr(108) & Chr(111) & Chr(103) & Chr(32) & Chr(105) & Chr(115) & Chr(32) & Chr(111) & Chr(110)
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(32) & Chr(33) & Chr(33) & Chr(33) & vbCrLf
newlog.SelColor = &H80&
End If
newlog.SetFocus
End Sub

Private Sub Command4_Click()
If Form1.Height = 3045 Then
Command3.Enabled = True
Command4.Enabled = False
Else
Form1.Height = Form1.Height - 1000
Command3.Enabled = True
Command4.Enabled = False
Text1.Visible = False
Label4.Caption = 2
newlog.SelColor = &H80000002
newlog.SelText = Chr(33) & Chr(33) & Chr(33) & Chr(32)
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, Chr(71) & Chr(101) & Chr(110) & Chr(101) & Chr(114) & Chr(97) & Chr(108) & Chr(32) & Chr(68) & Chr(97) & Chr(116) & Chr(101))
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & Chr(32) & Chr(45) & Chr(32) & Chr(83) & Chr(111) & Chr(102) & Chr(116) & Chr(32) & Chr(108) & Chr(111) & Chr(103) & Chr(32) & Chr(105) & Chr(115) & Chr(32) & Chr(111) & Chr(110)
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(32) & Chr(33) & Chr(33) & Chr(33) & vbCrLf
newlog.SelColor = &H80&
End If
newlog.SetFocus
End Sub

Private Sub Command5_Click()
Open Chr(76) & Chr(111) & Chr(103) & Chr(46) & Chr(116) & Chr(120) & Chr(116) For Output As #1
    Print #1, newlog.Text
Close #1
End Sub

Private Sub Command6_Click()
Form1.Enabled = False
Form2.Show
Form2.Enabled = True
End Sub

Private Sub Command7_Click()
For i = Winsock.LBound + 1 To Winsock.UBound
Winsock(i).Close
Unload Winsock(i)
Next i
Winsock(0).Close
End
End Sub

Private Sub Form_Load()
Text3.ToolTipText = Chr(80) & Chr(111) & Chr(114) & Chr(116) & Chr(115) & Chr(32) & Chr(97) & Chr(114) & Chr(101) & Chr(32) & Chr(111) & Chr(112) & Chr(101) & Chr(110) & Chr(101) & Chr(100) & Chr(32) & Chr(105) & Chr(110) & Chr(32) & Chr(111) & Chr(114) & Chr(100) & Chr(101) & Chr(114) & Chr(32) & Chr(111) & Chr(102) & Chr(32) & Chr(108) & Chr(101) & Chr(102) & Chr(116) & Chr(32) & Chr(116) & Chr(111) & Chr(32) & Chr(114) & Chr(105) & Chr(103) & Chr(104) & Chr(116) & Chr(46) & Chr(32) & Chr(83) & Chr(101) & Chr(112) & Chr(101) & Chr(114) & Chr(97) & Chr(116) & Chr(101) & Chr(32) & Chr(101) & Chr(97) & Chr(99) & Chr(104) & Chr(32) & Chr(112) & Chr(111) & Chr(114) & Chr(116) & Chr(32) & Chr(119) & Chr(105) & Chr(116) & Chr(104) & Chr(32) & Chr(116) & Chr(104) & Chr(101) & Chr(32) & Chr(44) & Chr(32) & Chr(107) & Chr(101) & Chr(121)
End Sub

Private Sub Form_Resize()
Label1.Top = Form1.Height - 640
End Sub

Private Sub Form_Terminate()
'Making shure our program closes and doesnt hang in memory
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Making shure our program closes and doesnt hang in memory
End
End Sub

Private Sub newlog_Change()
'When ever the log text box changes make shure it auto scrolls to the bottom line!
    On Error Resume Next
    newlog.SelLength = 0
    If Len(newlog.Text) > 0 Then
        If Right$(newlog.Text, 1) = vbCrLf Then
            newlog.SelStart = Len(newlog.Text) - 1
            Exit Sub
        End If
        newlog.SelStart = Len(newlog.Text)
    End If
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
    Dim Numbers As Integer
    Dim msg As String
    Numbers = KeyAscii


    If (((Numbers < 48 Or Numbers > 57) And Numbers <> 8) And Numbers <> 44) Then
        KeyAscii = 0
    End If
    'Only allowing numbers and the , key to be pressed in the port text box
End Sub

Private Sub Timer1_Timer()
'Timer1 ones job is to update our label1!
Label1.Caption = Format(Now, Chr(71) & Chr(101) & Chr(110) & Chr(101) & Chr(114) & Chr(97) & Chr(108) & Chr(32) & Chr(68) & Chr(97) & Chr(116) & Chr(101)) + Chr(32) & Chr(45) & Chr(32) + Winsock(Index).LocalIP + Chr(32) & Chr(111) & Chr(110) & Chr(32) + Winsock(Index).LocalHostName ' so every 1/1000th's of a second it updates the label to the new time + the IP of your computer and the host name of your computer!
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
'timer2's job is to disconnect the intruder from yur computer 2 seconds after he has connected!
Winsock(Label13.Caption).Close 'Disconnecting the intruder from your computer
Winsock(Label13.Caption).Listen 'Listening for another intruder to connect to your computer!
If Check2.Value = vbChecked Then MsgBox Label5.Caption + Chr(32) & Chr(104) & Chr(97) & Chr(115) & Chr(32) & Chr(98) & Chr(101) & Chr(101) & Chr(110) & Chr(32) & Chr(100) & Chr(105) & Chr(115) & Chr(99) & Chr(111) & Chr(110) & Chr(110) & Chr(101) & Chr(99) & Chr(116) & Chr(101) & Chr(100) & Chr(32) & Chr(102) & Chr(114) & Chr(111) & Chr(109) & Chr(32) & Chr(121) & Chr(111) & Chr(117) & Chr(114) & Chr(32) & Chr(99) & Chr(111) & Chr(109) & Chr(112) & Chr(117) & Chr(116) & Chr(101) & Chr(114) & Chr(33), vbSystemModal, Chr(65) & Chr(76) & Chr(69) & Chr(82) & Chr(84) & Chr(33) 'If check2 is checked then send me a msgbox saying that he is being disconnected!
Timer2.Enabled = False 'Disabling this because there is no more intruder!

End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Winsock(Label12.Caption).SendData vbCrLf
Winsock(Label12.Caption).SendData "#shell aborted"
Text1.Text = Text1.Text & Label10.Caption & " has been disconnected!"
Winsock(Label12.Caption).Close 'Disconnecting the intruder from your computer
Winsock(Label12.Caption).Listen 'Listening for another intruder to connect to your computer!
Timer3.Enabled = False 'Disabling this because there is no more intruder!
End Sub

Private Sub Winsock_Close(Index As Integer)
Winsock(Index).Close
Winsock(Index).Listen
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
'if there is a error. countinue on with the operation
'aka dont hold up the program.

'This is when someone trie's to connect weather we allow them to or not and what happens when they try!
Winsock(Index).Close 'Under winsock contrl if someone wants to connect you must close the port to alow them!
Winsock(Index).Accept requestID 'Accepting there request to connect!
DoEvents
'Adding the the textbox log that someone has connected
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(33) & Chr(33) & Chr(33) & Chr(32)
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & Chr(87) & Chr(97) & Chr(114) & Chr(110) & Chr(105) & Chr(110) & Chr(103) & Chr(32)
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, Chr(71) & Chr(101) & Chr(110) & Chr(101) & Chr(114) & Chr(97) & Chr(108) & Chr(32) & Chr(68) & Chr(97) & Chr(116) & Chr(101))
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(32) & Chr(45) & Chr(32)
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Winsock(Index).RemoteHostIP
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & Chr(58)
DoEvents
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Winsock(Index).Tag
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & Chr(32) & Chr(67) & Chr(111) & Chr(110) & Chr(110) & Chr(101) & Chr(99) & Chr(116) & Chr(105) & Chr(111) & Chr(110) & Chr(32) & Chr(65) & Chr(116) & Chr(116) & Chr(101) & Chr(109) & Chr(112) & Chr(116) & Chr(101) & Chr(100)
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & Chr(32) & Chr(33) & Chr(33) & Chr(33) & vbCrLf
DoEvents
If Winsock(Index).Tag = "23" Then
Label7.Caption = Winsock(Index).Tag
Label10.Caption = Winsock(Index).RemoteHostIP
Label12.Caption = Winsock(Index).Index
Else
Label9.Caption = Winsock(Index).Tag
Label11.Caption = Winsock(Index).RemoteHostIP
Label13.Caption = Winsock(Index).Index
End If
If Label4.Caption = 1 Then
Text1.Text = Text1.Text & Chr(67) & Chr(111) & Chr(109) & Chr(112) & Chr(117) & Chr(116) & Chr(101) & Chr(114) & Chr(32) & Winsock(Index).RemoteHostIP & Chr(32) & Chr(111) & Chr(110) & Chr(32) & Chr(112) & Chr(111) & Chr(114) & Chr(116) & Chr(32) & Winsock(Index).Tag & Chr(32) & Chr(104) & Chr(97) & Chr(115) & Chr(32) & Chr(99) & Chr(111) & Chr(110) & Chr(110) & Chr(101) & Chr(99) & Chr(116) & Chr(101) & Chr(100) & Chr(46) & vbCrLf
If Form2.Check3.Value = vbChecked Then Text1.Text = Text1.Text & "There hostname is - " + ResolveHostname(Winsock(Index).RemoteHostIP) & vbCrLf
If Form2.Check4.Value = vbChecked Then ImaPingJ00 (Winsock(Index).RemoteHostIP)
Else
DoEvents
End If
If Winsock(Index).Tag = "23" Then
If Form2.Check5.Value = vbChecked Then
Winsock(Index).SendData "[root@localhost /]"
Timer3.Interval = Form2.Text1.Text & "000"
Timer3.Enabled = True 'Enabling timer2 so we can disconnect him from our computer!
End If
Else
DoEvents
Timer2.Interval = Form2.Text3.Text & "000"
Timer2.Enabled = True 'Enabling timer2 so we can disconnect him from our computer!
If Form2.Check1.Value = vbChecked Then Winsock(Index).SendData Text2.Text 'If check1 is checked then send a message to the intruder... the message is text2.text
If Form2.Check2.Value = vbChecked Then MsgBox Chr(67) & Chr(111) & Chr(109) & Chr(112) & Chr(117) & Chr(116) & Chr(101) & Chr(114) & Chr(32) + Winsock(Index).RemoteHostIP + Chr(58) + Winsock(Index).Tag + Chr(32) & Chr(104) & Chr(97) & Chr(115) & Chr(32) & Chr(99) & Chr(111) & Chr(110) & Chr(110) & Chr(101) & Chr(99) & Chr(116) & Chr(101) & Chr(100) & Chr(33) + vbCrLf + Chr(68) & Chr(105) & Chr(115) & Chr(99) & Chr(111) & Chr(110) & Chr(110) & Chr(101) & Chr(99) & Chr(116) & Chr(105) & Chr(110) & Chr(103) & Chr(32) & Chr(116) & Chr(104) & Chr(101) & Chr(109) & Chr(32) & Chr(102) & Chr(114) & Chr(111) & Chr(109) & Chr(32) & Chr(121) & Chr(111) & Chr(117) & Chr(114) & Chr(32) & Chr(99) & Chr(111) & Chr(109) & Chr(112) & Chr(117) & Chr(116) & Chr(101) & Chr(114) & Chr(33), vbSystemModal, Chr(65) & Chr(76) & Chr(69) & Chr(82) & Chr(84) & Chr(33) 'If check2 is checked then send us a msgbox saying that the intruder *IP* has tried to connnect to us!
Text1.Text = Text1.Text & vbCrLf & Chr(73) & Chr(80) & Chr(32) & Winsock(Index).RemoteHostIP & Chr(32) & Chr(72) & Chr(97) & Chr(115) & Chr(32) & Chr(98) & Chr(101) & Chr(101) & Chr(110) & Chr(32) & Chr(100) & Chr(105) & Chr(115) & Chr(99) & Chr(111) & Chr(110) & Chr(110) & Chr(101) & Chr(99) & Chr(116) & Chr(101) & Chr(100) & Chr(33) & vbCrLf
DoEvents
End If
End Sub


Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'If the intruder trie's to send us any data
Dim incomingdata As String 'Dim the string as a string =)
Dim whatsup1 As String
Dim rawkon  As String
If Label4.Caption = 1 Then
'hard log is on
DoEvents
Winsock(Index).GetData incomingdata
'we are getting the incoming data
whatsup1 = Right$(incomingdata, 6)
If Form2.Check5.Value = vbChecked Then
If whatsup1 = vbCrLf Then
Winsock(Index).SendData Chr$(13)
Text1.Text = Text1.Text & Winsock(Index).RemoteHostIP & " - " & Label8.Caption & vbCrLf
Winsock(Index).SendData "[root@localhost /] " & Label8.Caption & vbCrLf
Winsock(Index).SendData "[root@localhost /]"
Label8.Caption = ""
DoEvents
Else
DoEvents
Label8.Caption = Label8.Caption & incomingdata
End If
End If
End If
End Sub
