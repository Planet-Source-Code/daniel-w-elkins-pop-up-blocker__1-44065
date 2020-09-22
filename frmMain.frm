VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E1272D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pop-Up Blocker"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E1272D&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0ECA
   ScaleHeight     =   3345
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmMain.frx":BD6C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4440
      MouseIcon       =   "frmMain.frx":C076
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.OptionButton optStart 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton optStop 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpStatus 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2708
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I would like to START this service."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   2940
   End
   Begin VB.Label lblStop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I would like to STOP this service."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   2835
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Wether You Want This Service Stopped or Started"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   970
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      Top             =   960
      Width           =   5655
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5760
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   120
      Picture         =   "frmMain.frx":C380
      Stretch         =   -1  'True
      Top             =   960
      Width           =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   5880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This simple utility either permits or blocks the net send messages channeled through the Messenger Service."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":E49E
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'These are the constants to tell the program wether the Messenger service is turned on or off.
Private Const ColorVal As Long = 14755629
Private Const SERVICE_STARTED As Integer = 1
Private Const SERVICE_STOPPED As Integer = 2

Dim iCurStatus As Integer

Private Sub cmdApply_Click()
If optStop.Value Then 'If the user has the "Stop service" option checked then...
shpStatus.FillColor = vbRed 'Change the shape's color to red to notify the user that the service is stopped.
iCurStatus = SERVICE_STOPPED 'Change the variable accordingly, so the program now knows that the service is stopped.
Call StopMsgService 'Stop the service.
Else 'If the user has the "Start service" option checked then...
shpStatus.FillColor = vbGreen 'The shape's color becomes green to notify the user that the service has started.
iCurStatus = SERVICE_STARTED 'Change the variable.
Call StartMsgService 'Start the service.
End If
Call SaveSettings 'Save the settings so the program knows the Service's status next time it loads.
End
End Sub

Private Sub cmdCancel_Click()
End 'Hmmm, wonder what this does!
End Sub

Private Sub Form_Load()
optStop.BackColor = ColorVal 'ColorVal is the color value of the background image, so the option boxes 'match'.
optStart.BackColor = ColorVal
Call ReadSettings 'Read the settings so the program knows the current status of the Messenger service.
If iCurStatus = SERVICE_STOPPED Then 'If the Service is currently stopped then...
shpStatus.FillColor = vbRed 'Change the shape's color.
optStop.Value = 1 'And check the "Stop service" option box for default.
ElseIf iCurStatus < 2 Then '<- Does the opposite.
shpStatus.FillColor = vbGreen
optStart.Value = 1
End If
End Sub

Private Sub lblStart_Click()
optStart.Value = 1
End Sub

Private Sub lblStop_Click()
optStop.Value = 1
End Sub

Private Sub StartMsgService()
Shell "net stop Messenger", vbHide 'Easy, huh ?
End Sub

Private Sub StopMsgService()
Shell "net start Messenger", vbHide
End Sub

Private Sub SaveSettings()
'Save the settings into the registry.
SaveSetting "StopMsg", "Main", "ServiceStatus", iCurStatus
End Sub

Private Sub ReadSettings()
'Read the settings from the registry. 2 (Service is stopped) is the default value for the Messenger service.
iCurStatus = Val(GetSetting("StopMsg", "Main", "ServiceStatus", 2))
End Sub
