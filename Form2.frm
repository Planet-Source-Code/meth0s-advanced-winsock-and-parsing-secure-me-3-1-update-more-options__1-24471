VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Secure me (+) 3.1 Options"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5910
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Close Options"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Disconnect times."
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   5655
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Text            =   "20"
         Top             =   220
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Second(s)."
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Disconnect all other ports in"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   525
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Second(s)."
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Disconnect connections on port 23 in"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Standerd Options"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "Send Intruder Message:"
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Text            =   "You do not have permission to access this service and are being reported to your ISP for malicious activatie."
         Top             =   480
         Width           =   3315
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Send me a visual alert when a intruder tries to connect to my computer."
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Advanced Options - Hard log only!"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5655
      Begin VB.CheckBox Check5 
         Caption         =   "If Telnet port is open. Fake linux server with root@local"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   5055
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Ping there IP when they connect and find out there delay."
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4935
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Resolve Hostname using there IP when they connect."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   5055
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Enabled = True
Form2.Enabled = False
Form2.Visible = False
End Sub

Private Sub Form_Terminate()
Form1.Enabled = True
Form2.Enabled = False
Form2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
Form2.Enabled = False
Form2.Visible = False
End Sub

