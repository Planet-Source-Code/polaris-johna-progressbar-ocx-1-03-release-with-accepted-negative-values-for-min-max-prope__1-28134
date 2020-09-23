VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Johna ProgressBar Control example"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin Projet2.Johna_BAR BAR3 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      Value           =   60
      Bar_COLOR       =   98979694
   End
   Begin Projet2.Johna_BAR BAR2 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      Value           =   58
      Border_Type     =   1
      Bar_COLOR       =   44485051
   End
   Begin Projet2.Johna_BAR BAR1 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      MAX             =   150
      Value           =   111
      Bar_COLOR       =   60616263
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test johna progress bar"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "FX border  type 3"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "FX border  type 2"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "FX border  type 1"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Johna_BAR1_GotFocus()

End Sub

Private Sub Command1_Click()
Dim I
Randomize Timer
BAR1.Orientation = Johna_Horizontal

'BAR1.Bar_COLOR = &HFDFDFF1

 
For I = 0 To 100 Step 0.1
  BAR1.Value = I
  BAR2.Value = I
  BAR3.Value = I
  DoEvents
Next I



End Sub

