VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   Caption         =   "Gravity Tutorial by tHa_imaX - DDC ELITE"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   572
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   790
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox chkGL 
      Caption         =   "Show Gravity Direction"
      Height          =   255
      Left            =   6840
      TabIndex        =   43
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CheckBox chSky 
      Caption         =   "Allow Sky?"
      Height          =   270
      Left            =   5760
      TabIndex        =   41
      Top             =   8280
      Value           =   1  'Aktiviert
      Width           =   1020
   End
   Begin VB.Frame Frame4 
      Caption         =   "Nature"
      Height          =   1695
      Left            =   6945
      TabIndex        =   36
      Top             =   6600
      Width           =   1710
      Begin MSComctlLib.Slider slWind 
         Height          =   270
         Left            =   30
         TabIndex        =   38
         Top             =   555
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Max             =   100
      End
      Begin MSComctlLib.Slider slGravity 
         Height          =   270
         Left            =   45
         TabIndex        =   40
         Top             =   1275
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Max             =   100
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         Caption         =   "Gravity 1\1000"
         Height          =   270
         Left            =   90
         TabIndex        =   39
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         Caption         =   "Wind x\100"
         Height          =   270
         Left            =   75
         TabIndex        =   37
         Top             =   270
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkCLS 
      Caption         =   "Do CLS?"
      Height          =   285
      Left            =   4800
      TabIndex        =   35
      Top             =   8280
      Value           =   1  'Aktiviert
      Width           =   960
   End
   Begin VB.Frame Frame3 
      Caption         =   "Info"
      Height          =   6630
      Left            =   10080
      TabIndex        =   22
      Top             =   0
      Width           =   1695
      Begin VB.CheckBox chintFS 
         Caption         =   "int FS Values?"
         Height          =   255
         Left            =   150
         TabIndex        =   33
         Top             =   2220
         Width           =   1365
      End
      Begin VB.Label nfoRX 
         Caption         =   "RealX: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label nfoRY 
         Caption         =   "RealY: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   870
         Width           =   1455
      End
      Begin VB.Label nfoB 
         Caption         =   "Ball is out of range"
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   120
         TabIndex        =   30
         Top             =   3225
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label nfofY 
         Caption         =   "fSy:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1965
         Width           =   1455
      End
      Begin VB.Label nfofX 
         Caption         =   "fSx:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1725
         Width           =   1455
      End
      Begin VB.Label nfoColor 
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   2895
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Color"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2895
         Width           =   375
      End
      Begin VB.Label nfoG 
         Caption         =   "Mass: 0kg"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2655
         Width           =   1455
      End
      Begin VB.Label nfoY 
         Caption         =   "VB  Y: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label nfoX 
         Caption         =   "VB  X: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1140
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add a Ball"
      Height          =   1695
      Left            =   3855
      TabIndex        =   8
      Top             =   6600
      Width           =   3015
      Begin VB.CheckBox chkRV 
         Caption         =   "rnd values"
         Height          =   300
         Left            =   105
         TabIndex        =   34
         ToolTipText     =   "if you select this you'll add a random point if you click on <ADD>"
         Top             =   1290
         Width           =   1080
      End
      Begin VB.TextBox txtG 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         Height          =   315
         Left            =   2400
         TabIndex        =   19
         Text            =   "1,7"
         Top             =   585
         Width           =   495
      End
      Begin VB.CommandButton cmdADD 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1245
         TabIndex        =   17
         ToolTipText     =   "click here to add a point in to our system"
         Top             =   1320
         Width           =   1635
      End
      Begin VB.CommandButton Command1 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         ToolTipText     =   "click here to randomize the values."
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtC 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         Height          =   300
         Left            =   600
         TabIndex        =   14
         Text            =   "0000"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtY 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         Height          =   315
         Left            =   1335
         TabIndex        =   13
         Text            =   "10"
         Top             =   585
         Width           =   615
      End
      Begin VB.TextBox txtX 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Text            =   "400"
         Top             =   585
         Width           =   615
      End
      Begin VB.TextBox txtNBN 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "4"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   585
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.CheckBox chkPntPic 
      Caption         =   "Do PaintPicture (expert only)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2400
      TabIndex        =   7
      Top             =   8280
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ball Config"
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   6600
      Width           =   3735
      Begin VB.CommandButton Command3 
         Caption         =   "Random him"
         Height          =   255
         Left            =   2580
         TabIndex        =   42
         Top             =   600
         Width           =   1020
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove him"
         Height          =   255
         Left            =   1395
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin MSComctlLib.Slider slPowerY 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "kick power of our Y"
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   1
         Value           =   1
      End
      Begin VB.CommandButton cmdKick 
         Caption         =   "Kick him"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbBall 
         Appearance      =   0  '2D
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Text            =   "Which Ball?"
         Top             =   240
         Width           =   3495
      End
      Begin MSComctlLib.Slider slPowerX 
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Min             =   -15
         Max             =   15
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         Caption         =   "Power (for kicking only) Y - X"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   3255
      End
   End
   Begin VB.PictureBox rDraw 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6480
      Left            =   0
      ScaleHeight     =   428
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   658
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   9930
   End
   Begin VB.Timer tGrav 
      Interval        =   1
      Left            =   11400
      Top             =   8160
   End
   Begin VB.PictureBox pDraw 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6480
      Left            =   0
      ScaleHeight     =   428
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   658
      TabIndex        =   0
      Top             =   120
      Width           =   9930
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Gravity Tutorial by tHa_imaX [Digital Death Crew 2oo4]
' Uploaded to Planet S0urCe c0de
' Any Questions? Contact me!
' icQ  . 82186397
' emaiL. the_imax@yahoo.de
' Pure .vB. Code

Private Sub chkPntPic_Click()
    If chkPntPic Then
        rDraw.Visible = True
    Else
        rDraw.Visible = False
    End If
End Sub



Private Sub cmbBall_Click()
   If cmbBall <> "ALL" Then
    nfoG = "Mass: " & Pnts(cmbBall).g & "kg"
    nfoColor.BackColor = Pnts(cmbBall).Col
   End If
End Sub

Private Sub cmdADD_Click()
    If chkRV Then Command1_Click
    Pnts.Add New cPoint, txtNBN
    SetPoint txtNBN, Pnts.Count, txtX, txtY, txtG, Val(txtC)
    SetControls
    txtNBN = Pnts.Count + 1
End Sub

Private Sub cmdKick_Click()
    On Error Resume Next
    If cmbBall = "ALL" Then
     For i = 1 To Pnts.Count
      Pnts(i).fSy = slPowerY
     Pnts(i).fSx = slPowerX
     Next
    Else
     Pnts(cmbBall).fSy = slPowerY * -1
     Pnts(cmbBall).fSx = slPowerX * -1
    End If
End Sub

Private Sub Command1_Click()
Randomize Timer
Dim r, g, b
r = 1 + Rnd * 255
g = 1 + Rnd * 255
b = 1 + Rnd * 255
txtC = RGB(r, g, b)
txtG = 2 + Rnd * 4
txtX = 10 + Rnd * pDraw.ScaleWidth - 10
End Sub

Private Sub Command2_Click()
    If cmbBall = "ALL" Then
     Do Until Pnts.Count = 0
        Pnts.Remove 1
     Loop
    Else
     If cmbBall.ListCount > 1 Then Pnts.Remove cmbBall
    End If
    SetControls
    txtNBN = Pnts.Count + 1
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    If cmbBall = "ALL" Then
     For i = 1 To Pnts.Count
      Dim tmp As Integer
      Dim tmp2 As Integer
      Do Until tmp <> 0 And tmp2 <> 0
        tmp = Int(-1 + Rnd * 1)
        tmp2 = Int(-1 + Rnd * 1)
      Loop
      Pnts(i).fSy = (1 + Rnd * 100) * tmp
      Pnts(i).fSx = (1 + Rnd * 100) * tmp2
     Next
    Else
      Do Until tmp <> 0 And tmp2 <> 0
        tmp = Int(-1 + Rnd * 1)
        tmp2 = Int(-1 + Rnd * 1)
      Loop
       Pnts(cmbBall).fSy = (1 + Rnd * 100) * tmp
       Pnts(cmbBall).fSx = (1 + Rnd * 100) * tmp2
    End If
End Sub

Private Sub Form_Load()
    Randomize Timer
    Pnts.Add New cPoint, "1"
    SetPoint "1", 1, 50, 1, 1.5, vbBlack
    
    Pnts.Add New cPoint, "2"
    SetPoint "2", 2, 150, 1, 2, vbRed
    
    Pnts.Add New cPoint, "3"
    SetPoint "3", 3, 200, 1, 2.2, vbCyan
    
    Wind = 0
    vGrav = 0.05
    SetControls

End Sub

Private Sub slGravity_Change()
    vGrav = slGravity.Value / 100
End Sub

Private Sub slWind_Change()
    Wind = slWind.Value / 100
End Sub

Private Sub tGrav_Timer()
 If chkCLS Then pDraw.Cls
 For i = 1 To Pnts.Count
  With Pnts(i)
  
  .fSy = .fSy + ((vGrav * .g) * 1)
  
  If .Y > pDraw.ScaleHeight - 5 And .fSy > 0 Then
    .fSy = (.fSy * -1) / (.g + vGrav / 2) + (0.1 + Rnd * 0.5)
  End If
  
  
  If chSky = False Then
    If .Y <= 0 Then
        .Y = 5
        .fSy = (.fSy * -1) / (.g + vGrav / 2) ' + (0.1 + Rnd * 0.5)
    End If
  End If
  
  If .Y > pDraw.ScaleHeight - 5 And vGrav > 0 Then .Y = pDraw.ScaleHeight - 5 'Callibrate to Floor
    
  If .fSy > -1 And .fSy < 1 And .Y = pDraw.ScaleHeight - 5 Then 'Collision with Floor
    .fSy = 0
  Else
    .Y = .Y + .fSy
  End If
  
  
  
  .fSx = Wind + .fSx
  If .X > pDraw.ScaleWidth Then .X = pDraw.ScaleWidth - 5
  If .X < 0 Then .X = 5
  
  
  If .X = pDraw.ScaleWidth - 5 Or .X = 5 Then
    'If .X < 5 Then .X = 5
    .fSx = (.fSx * -1) / (.g + vGrav / 2) ' + (0.1 + Rnd * 0.5)
  End If
  
  If .fSy <= 5 Then
    If .Y = pDraw.ScaleHeight - 5 Then
        .fSx = (.fSx / 100) * (100 - .g)
    End If
  End If
  
  'If .fSx > -0.3 And .fSx < 0.3 Then
  '  .fSx = 0
  'Else
    .X = .X + .fSx
  'End If
  'Debug.Print "y: " & .Y & " @fsy: " & .fsy
  pDraw.DrawWidth = 1
  If chkGL And .fSy <> 0 Then
    If .fSy > 0 Then
        pDraw.PSet (.X, .Y + 5)
    Else
        pDraw.PSet (.X, .Y - 5)
    End If
    
  End If
  
  If i = cmbBall Then
    pDraw.Circle (.X, .Y), 6, vbRed
  End If
  
  pDraw.DrawWidth = 5
  pDraw.PSet (.X, .Y), .Col
  
  RefreshNFO
  If chkPntPic Then rDraw.PaintPicture pDraw.Image, 0, 0
  End With
  
 
 Next
 
End Sub
