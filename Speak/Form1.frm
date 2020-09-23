VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   3780
   ClientTop       =   2055
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   3525
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   -360
      Top             =   3120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Speak"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS Speak 
      Height          =   1215
      Left            =   960
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4050
      Left            =   -960
      Picture         =   "Form1.frx":0058
      Top             =   -120
      Width           =   5400
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
' Software by Sahib
'****************************************************
Dim A(10) As String         'A global array to store text
Dim I As Integer
Private Sub Command1_Click()
    Speak.Speak Text1.Text      'Speak the text typed in Text box
End Sub

Private Sub Form_Load()
    A(1) = "HELLO HOW ARE YOU"
    A(2) = "YOU HAVE NOT ANSWERED ME"
    A(3) = "GOOD SMILE"
    A(4) = "YOU ARE LOOKING VERY SMART TODAY"
    A(5) = "ARE YOU GOING TO MEET SOMEONE SPECIAL TODAY"
    A(6) = "IF YOU WEAR WHITE SHIRT AND BROWN TROUSER IT WILL MAKE YOU MORE SMART"
    Speak.MaxSpeed = 40
End Sub

Private Sub Timer1_Timer()
    I = I + 1
    If I > 6 Then   'Timer to move to array elements
        Timer1.Enabled = False
    End If
    Label1.Caption = A(I)
    Speak.Speak Label1.Caption
    
End Sub
