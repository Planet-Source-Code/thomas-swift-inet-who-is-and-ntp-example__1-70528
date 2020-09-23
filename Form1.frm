VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1
Caption         =   "Inet Who Is And NTP Example"
ClientHeight    =   5355
ClientLeft      =   60
ClientTop       =   450
ClientWidth     =   6300
LinkTopic       =   "Form1"
ScaleHeight     =   5355
ScaleWidth      =   6300
StartUpPosition =   3 'Windows Default
Begin VB.CommandButton Command1
Caption         =   "Go"
Height          =   270
Left            =   3885
TabIndex        =   2
Top             =   97
Width           =   405
End
Begin VB.TextBox IP
Height          =   285
Left            =   2010
TabIndex        =   1
Text            =   "66.79.189.66"
Top             =   90
Width           =   1800
End
Begin VB.TextBox Text1
Height          =   4470
Left            =   30
MultiLine       =   -1 'True
ScrollBars      =   2 'Vertical
TabIndex        =   0
Top             =   420
Width           =   6240
End
Begin InetCtlsObjects.Inet Inet1
Left            =   30
Top             =   15
_ExtentX        =   1005
_ExtentY        =   1005
_Version        =   393216
AccessType      =   1
End
Begin VB.Label Label1
Alignment       =   2 'Center
BackColor       =   &H00000000&
BeginProperty Font
Name            =   "Arial"
Size            =   14.25
Charset         =   0
Weight          =   700
Underline       =   0 'False
Italic          =   0 'False
Strikethrough   =   0 'False
EndProperty
ForeColor       =   &H000000FF&
Height          =   390
Left            =   345
TabIndex        =   3
Top             =   4920
Width           =   5625
End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By:Thomas A. Swift May 12,2008
Option Explicit
Private Sub Command1_Click()
    'On Error Resume Next
    Text1.SetFocus
    Text1.Text = WhoIs(Inet1, IP.Text)
    Label1.Caption = "UTC Date/Time Is: " & GetTime(Inet1)
    'Debug.Print Inet1.GetHeader '"HTTP/1.0 200 OK"
End Sub

