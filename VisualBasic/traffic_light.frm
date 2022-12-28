VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21405
   FillColor       =   &H0080FFFF&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   17.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "traffic_light.frx":0000
   ScaleHeight     =   10095
   ScaleWidth      =   21405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Timer_Green 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15960
      TabIndex        =   31
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox Timer_Red 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15960
      TabIndex        =   30
      Top             =   5400
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   14040
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Exit_button 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Timer Timer12 
      Interval        =   1000
      Left            =   18600
      Top             =   9240
   End
   Begin VB.CommandButton Realtime_button 
      BackColor       =   &H00FF8080&
      Caption         =   "REAL TIME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   19920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox SetT_Night_mode 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   19920
      TabIndex        =   17
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox SetT_Auto_mode 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   19920
      TabIndex        =   16
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Timer Timer11 
      Interval        =   1000
      Left            =   18000
      Top             =   9240
   End
   Begin VB.Timer Timer10 
      Left            =   16800
      Top             =   9840
   End
   Begin VB.Timer Timer9 
      Left            =   16080
      Top             =   9840
   End
   Begin VB.Timer Timer8 
      Left            =   15360
      Top             =   9840
   End
   Begin VB.Timer Timer7 
      Left            =   14640
      Top             =   9840
   End
   Begin VB.CommandButton R1G2_button 
      BackColor       =   &H00FFFFC0&
      Caption         =   "R1 - G2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   17880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton R2G1_button 
      BackColor       =   &H00FFFFC0&
      Caption         =   "R2 - G1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   17880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Night_button 
      BackColor       =   &H00FF80FF&
      Caption         =   "NIGHT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Manual_button 
      BackColor       =   &H00FFFFC0&
      Caption         =   "MANUAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   17880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Auto_button 
      BackColor       =   &H00C0FFC0&
      Caption         =   "AUTO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Stop_button 
      BackColor       =   &H0000FFFF&
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Start_button 
      BackColor       =   &H0000FF00&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Timer Timer6 
      Left            =   13920
      Top             =   9840
   End
   Begin VB.Timer Timer5 
      Left            =   16800
      Top             =   9240
   End
   Begin VB.Timer Timer4 
      Left            =   16080
      Top             =   9240
   End
   Begin VB.Timer Timer3 
      Left            =   15360
      Top             =   9240
   End
   Begin VB.Timer Timer2 
      Left            =   14640
      Top             =   9240
   End
   Begin VB.Timer Timer1 
      Left            =   13920
      Top             =   9240
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TIME GREEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   15960
      TabIndex        =   33
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "TIME RED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   15840
      TabIndex        =   32
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "IUH - FEET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label13 
      Caption         =   "Ho va ten: Pham Chi Hieu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label12 
      Caption         =   "MSSV: 19433791"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4440
      TabIndex        =   27
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "Lop: DHDKTD15A_Nhom 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label10 
      Caption         =   "STT: 08"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4440
      TabIndex        =   25
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Date 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13680
      TabIndex        =   23
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   21360
      TabIndex        =   22
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   21360
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SET TIME AUTO MODE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   19920
      TabIndex        =   19
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SET TIME NIGHT MODE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   19920
      TabIndex        =   18
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   6600
      TabIndex        =   15
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   9000
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Lane2 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Index           =   1
      Left            =   5880
      TabIndex        =   13
      Top             =   3000
      Width           =   975
   End
   Begin VB.Shape Yellow2 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Green2 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Red2 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   615
   End
   Begin VB.Shape Green1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   615
   End
   Begin VB.Shape Red1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Lane1 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Index           =   1
      Left            =   6960
      TabIndex        =   12
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Yellow1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Lane2 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Index           =   0
      Left            =   13200
      TabIndex        =   11
      Top             =   6840
      Width           =   975
   End
   Begin VB.Shape Yellow1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   12480
      TabIndex        =   10
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   10080
      TabIndex        =   9
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label Lane1 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Index           =   0
      Left            =   12120
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "TRAFFIC LIGHT SIMULATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   7560
      TabIndex        =   0
      Top             =   360
      Width           =   5175
   End
   Begin VB.Shape Red2 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   8640
      Width           =   615
   End
   Begin VB.Shape Green2 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   615
   End
   Begin VB.Shape Yellow2 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   615
   End
   Begin VB.Shape Green1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   615
   End
   Begin VB.Shape Red1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   14040
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t1, t2, t3, t4 As Integer
Dim i As Integer
Dim opt As Boolean
Dim x1, x2 As Integer
Dim s As Integer

Private Sub Exit_button_Click()
    End
End Sub

Private Sub Form_Load()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    Timer7.Enabled = False
    Timer8.Enabled = False
    Timer9.Enabled = False
    Timer10.Enabled = False
    Timer11.Enabled = False
    Timer12.Enabled = True
    
    Auto_button.Enabled = True
    Night_button.Enabled = True
    Manual_button.Enabled = True
    Start_button.Enabled = False
    Stop_button.Enabled = True
    R1G2_button.Enabled = False
    R2G1_button.Enabled = False
    Realtime_button = False
    
    MSComm1.CommPort = 5
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.PortOpen = True
    
End Sub

'Auto button
Private Sub Auto_button_Click()

    'MSComm1.Output = "a"
    
    opt = True
    Start_button.Enabled = True
    Stop_button.Enabled = True
    Night_button.Enabled = False
    Manual_button.Enabled = False
    Auto_button.Enabled = False
End Sub

'Manual button
Private Sub Manual_button_Click()
    Start_button.Enabled = False
    Stop_button.Enabled = True
    Auto_button.Enabled = False
    Night_button.Enabled = False
    Manual_button.Enabled = False
    R1G2_button.Enabled = True
    R2G1_button.Enabled = True
End Sub

'Night button
Private Sub Night_button_Click()

    'MSComm1.Output = "c"
    
    opt = False
    Start_button.Enabled = True
    Stop_button.Enabled = True
    Auto_button.Enabled = False
    Manual_button.Enabled = False
    Night_button.Enabled = False
    R1G2_button.Enabled = False
    R2G1_button.Enabled = False
End Sub

Private Sub R1G2_button_Click()
    Red1(0).FillColor = vbRed
    Red1(1).FillColor = vbRed
    Green2(0).FillColor = vbGreen
    Green2(1).FillColor = vbGreen
    
    Red2(0).FillColor = vbBlack
    Red2(1).FillColor = vbBlack
    Green1(0).FillColor = vbBlack
    Green1(1).FillColor = vbBlack
    
    MSComm1.Output = "g"
End Sub

Private Sub R2G1_button_Click()
    Red2(0).FillColor = vbRed
    Red2(1).FillColor = vbRed
    Green1(0).FillColor = vbGreen
    Green1(1).FillColor = vbGreen
    
    Red1(0).FillColor = vbBlack
    Red1(1).FillColor = vbBlack
    Green2(0).FillColor = vbBlack
    Green2(1).FillColor = vbBlack
    
    MSComm1.Output = "h"
End Sub

Private Sub Realtime_button_Click()
    Timer11.Enabled = True
    
    Auto_button.Enabled = False
    Night_button.Enabled = False
    Manual_button.Enabled = False
    R1G2_button.Enabled = False
    R2G1_button.Enabled = False
End Sub

'Private Sub SetT_Auto_mode_Change()
    'Realtime_button.Enabled = True
'End Sub

'Private Sub SetT_Night_mode_Change()
    'Realtime_button.Enabled = True
'End Sub

'Chuong trinh nut Start
Private Sub Start_button_Click()

    'Set thoi gian
    x1 = Val(Timer_Red.Text) * 1000
    x2 = Val(Timer_Green.Text) * 1000

    If opt = True Then
        MSComm1.Output = "d"
        Red1(0).FillColor = vbRed
        Red1(1).FillColor = vbRed
        
        Green2(0).FillColor = vbGreen
        Green2(1).FillColor = vbGreen
        
        Timer1.Enabled = True
        Timer1.Interval = x2
        
        t1 = x1 / 1000
        t2 = x2 / 1000
        t3 = x1 / 1000
        t4 = x2 / 1000
        
        'MSComm1.Output = "a"
        
        Timer5.Enabled = True
        Timer5.Interval = 1000
        Timer6.Enabled = True
        Timer6.Interval = 1000
        Timer7.Enabled = True
        Timer7.Interval = 1000
        Timer8.Enabled = True
        Timer8.Interval = 1000
    End If
    
    If opt = False Then
        Timer9.Enabled = True
        Timer9.Interval = 1000
        Timer10.Enabled = True
        Timer10.Interval = 1000
    End If
End Sub

Private Sub Stop_button_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    Timer7.Enabled = False
    Timer8.Enabled = False
    Timer9.Enabled = False
    Timer10.Enabled = False
    
    Red1(0).FillColor = vbBlack
    Red1(1).FillColor = vbBlack
    Red2(0).FillColor = vbBlack
    Red2(1).FillColor = vbBlack
    
    Green1(0).FillColor = vbBlack
    Green1(1).FillColor = vbBlack
    Green2(0).FillColor = vbBlack
    Green2(1).FillColor = vbBlack
    
    Yellow1(0).FillColor = vbBlack
    Yellow1(1).FillColor = vbBlack
    Yellow2(0).FillColor = vbBlack
    Yellow2(1).FillColor = vbBlack
    
    Lane1(0).Caption = 0
    Lane1(1).Caption = 0
    Lane2(0).Caption = 0
    Lane2(1).Caption = 0
    
    Auto_button.Enabled = True
    Night_button.Enabled = True
    Manual_button.Enabled = True
    
    MSComm1.Output = "z"
End Sub

'Timer Auto mode
Private Sub Timer1_Timer()

    MSComm1.Output = "a"

    Timer2.Enabled = True
    Timer2.Interval = x1 - x2
    
    Yellow2(0).FillColor = vbYellow
    Yellow2(1).FillColor = vbYellow
    
    Red1(0).FillColor = vbRed
    Red1(1).FillColor = vbRed
    
    Green2(0).FillColor = vbBlack
    Green2(1).FillColor = vbBlack
    
    'MSComm1.Output = Chr(20)
    
    t2 = (x1 - x2) / 1000
    t4 = (x1 - x2) / 1000
    
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    
    MSComm1.Output = "b"
    
    Red2(0).FillColor = vbRed
    Red2(1).FillColor = vbRed
    
    Green1(0).FillColor = vbGreen
    Green1(1).FillColor = vbGreen
    
    Red1(0).FillColor = vbBlack
    Red1(1).FillColor = vbBlack
    
    Yellow2(0).FillColor = vbBlack
    Yellow2(1).FillColor = vbBlack
    
    'MSComm1.Output = Chr(33)
    
    Timer3.Enabled = True
    Timer3.Interval = x2
    Timer2.Enabled = False
    
    t2 = x1 / 1000
    t1 = x2 / 1000
    t4 = x1 / 1000
    t3 = x2 / 1000
End Sub

Private Sub Timer3_Timer()

    MSComm1.Output = "c"
    
    Red2(0).FillColor = vbRed
    Red2(1).FillColor = vbRed
    
    Yellow1(0).FillColor = vbYellow
    Yellow1(1).FillColor = vbYellow
    
    Green1(0).FillColor = vbBlack
    Green1(1).FillColor = vbBlack
    
    'MSComm1.Output = Chr(34)
    
    Timer4.Enabled = True
    Timer4.Interval = x1 - x2
    Timer3.Enabled = False
    
    t1 = (x1 - x2) / 1000
    t3 = (x1 - x2) / 1000
End Sub

Private Sub Timer4_Timer()

    MSComm1.Output = "d"

    Red1(0).FillColor = vbRed
    Red1(1).FillColor = vbRed
    
    Green2(0).FillColor = vbGreen
    Green2(1).FillColor = vbGreen
    
    Red2(0).FillColor = vbBlack
    Red2(1).FillColor = vbBlack
    
    Yellow1(0).FillColor = vbBlack
    Yellow1(1).FillColor = vbBlack
    
    MSComm1.Output = Chr(12)
    
    Timer1.Enabled = True
    Timer1.Interval = x2
    Timer4.Enabled = False
    
    t1 = x1 / 1000
    t2 = x2 / 1000
    t3 = x1 / 1000
    t4 = x2 / 1000
End Sub

'Chuong trinh hien thi thoi gian dem nguoc o Lane 1
Private Sub Timer5_Timer()
    If t1 > 0 Then
    t1 = t1 - 1
    Lane1(0).Caption = t1
    End If
End Sub

'Chuong trinh hien thi thoi gian dem nguoc o Lane 2
Private Sub Timer6_Timer()
    If t2 > 0 Then
        t2 = t2 - 1
        Lane2(0).Caption = t2
    End If
End Sub

Private Sub Timer7_Timer()
    If t3 > 0 Then
        t3 = t3 - 1
        Lane1(1).Caption = t3
    End If
End Sub

Private Sub Timer8_Timer()
    If t4 > 0 Then
        t4 = t4 - 1
        Lane2(1).Caption = t4
    End If
End Sub

'Flashing for Night mode
Private Sub Timer9_Timer()

    MSComm1.Output = "e"

    Yellow1(0).FillColor = vbBlack
    Yellow1(1).FillColor = vbBlack
    Yellow2(0).FillColor = vbBlack
    Yellow2(1).FillColor = vbBlack
    
    'MSComm1.Output = Chr(18)
End Sub

'Flashing for Night mode
Private Sub Timer10_Timer()

    'MSComm1.Output = "f"
    
    Yellow1(0).FillColor = vbYellow
    Yellow1(1).FillColor = vbYellow
    Yellow2(0).FillColor = vbYellow
    Yellow2(1).FillColor = vbYellow
    
    'MSComm1.Output = Chr(0)
End Sub

'Timer Real Time
Private Sub Timer11_Timer()
    MSComm1.Output = "z"

    s = Second(Now)
    
    If s = Val(SetT_Night_mode.Text) Then
        Red1(0).FillColor = vbBlack
        Red1(1).FillColor = vbBlack
        Green2(0).FillColor = vbBlack
        Green2(1).FillColor = vbBlack
    
        Red2(0).FillColor = vbBlack
        Red2(1).FillColor = vbBlack
        Green1(0).FillColor = vbBlack
        Green1(1).FillColor = vbBlack
        
        Lane1(0).Caption = 0
        Lane1(1).Caption = 0
        Lane2(0).Caption = 0
        Lane2(1).Caption = 0
    
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
    
        Timer9.Enabled = True
        Timer9.Interval = 1000
        Timer10.Enabled = True
        Timer10.Interval = 1000
    End If
    If s = Val(SetT_Auto_mode.Text) Then
        Yellow1(0) = vbBlack
        Yellow1(1) = vbBlack
        Yellow2(0) = vbBlack
        Yellow2(1) = vbBlack
        
        Timer9.Enabled = False
        Timer10.Enabled = False
        
        opt = True
        Start_button_Click
    End If
End Sub

Private Sub Timer12_Timer()
    Date.Caption = Now
End Sub
