VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "3d"
   ClientHeight    =   4695
   ClientLeft      =   2670
   ClientTop       =   1500
   ClientWidth     =   8040
   FillColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0152
   ScaleHeight     =   4695
   ScaleWidth      =   8040
   Begin MSComDlg.CommonDialog cdicolor 
      Left            =   6960
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7560
      Top             =   2640
   End
   Begin VB.CommandButton Command4 
      Caption         =   "size"
      Height          =   255
      Left            =   1320
      TabIndex        =   30
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ok"
      Height          =   450
      Left            =   3480
      TabIndex        =   29
      Top             =   2280
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1320
      TabIndex        =   28
      Top             =   2280
      Width           =   2175
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   4560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   27
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   26
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   25
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H0080FF80&
      Height          =   255
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   24
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   23
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   22
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   4560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   21
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FF80FF&
      Height          =   255
      Left            =   4560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   20
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   4320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   3840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C000&
      Height          =   255
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox FX 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   0
      ScaleHeight     =   1650
      ScaleWidth      =   8160
      TabIndex        =   8
      Top             =   0
      Width           =   8160
   End
   Begin VB.TextBox txtMessage 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      IMEMode         =   3  'DISABLE
      Left            =   150
      MaxLength       =   20
      TabIndex        =   7
      Top             =   3900
      Width           =   7515
   End
   Begin VB.CommandButton high 
      Caption         =   "high"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton medium 
      Caption         =   "medium"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton low 
      Caption         =   "low"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox color 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   5880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BackGround"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Picture"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin a3d.Transparent Transparent1 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      MaskColor       =   255
   End
   Begin MSComDlg.CommonDialog cdisave 
      Left            =   6960
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdiopen 
      Left            =   6960
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label to1 
      Caption         =   "200"
      Height          =   255
      Left            =   7440
      TabIndex        =   32
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   405
      Left            =   6120
      Picture         =   "Form1.frx":80A38
      Top             =   1720
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   5520
      Picture         =   "Form1.frx":812EA
      Top             =   1725
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   5540
      Picture         =   "Form1.frx":81E6C
      Top             =   1725
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   6120
      Picture         =   "Form1.frx":829EE
      Top             =   1720
      Width           =   390
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   3960
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label sim 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2520
      TabIndex        =   31
      Top             =   2780
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "3DText"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "COLOR"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "BackGround"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "COLOR"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label d3 
      Caption         =   "10"
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label d2 
      Caption         =   "-20"
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label d1 
      Caption         =   "20"
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   1800
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intCount As Integer
Dim xrad As String, yrad, fromLeft, fromTop, pSpeed, PiniAngle, PiniAngleTemp, numOfRotation
Dim moving As Boolean, makeChangesToRegistry, settingsAltered, settings_saved
Dim lX As Single, lY As Single

Private Sub Command3_Click()
If List1.ListCount <> 0 Then
        FX.Font = List1.Text
    Else
        MsgBox "you have To choose the fonts first"
    End If
End Sub

Private Sub Command4_Click()
 If List1.ListCount <> 0 Then
Dim Size As Single
        Size = Val(InputBox("Enter the font size"))
        FX.FontSize = Val(Size)
    Else
        MsgBox "you have To choose the fonts first"
    End If
End Sub
Private Sub Form_Load()
    Dim NUM As Single
    Dim X As Single

    NUM = Screen.FontCount

    For X = 1 To NUM
        List1.AddItem Screen.Fonts(X)
    Next X
    List1.RemoveItem (0)
If LCase(Command) = "aniandexit" Then
    Form1.Visible = False
Else
    Show
    Set Transparent1.MaskPicture = Form1.Picture
End If
Image7.Picture = Form1.Icon
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moving = True
    lX = X
    lY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If moving Then
        Me.Move (Me.Left + X - lX), (Me.Top + Y - lY)
        DoEvents
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moving = False
End Sub
Private Sub color_Click()
cdicolor.ShowColor
FX.BackColor = cdicolor.color
color.BackColor = cdicolor.color
FX.Cls
For intCount = 200 To 250
    FX.ForeColor = RGB(intCount + d1.Caption, intCount + d2.Caption, intCount + d3.Caption)
    FX.CurrentX = intCount
    FX.CurrentY = intCount
    FX.Print txtMessage.Text
Next intCount
End Sub

Private Sub Command1_Click()
On Error GoTo l
   cdiopen.Filter = "(*.bmp)|*.bmp"
   cdiopen.ShowOpen
   FX.Picture = LoadPicture(cdiopen.FileName)
FX.Cls
For intCount = 200 To 250
    FX.ForeColor = RGB(intCount + d1.Caption, intCount + d2.Caption, intCount + d3.Caption)
    FX.CurrentX = intCount
    FX.CurrentY = intCount
    FX.Print txtMessage.Text
Next intCount
l:
End Sub

Private Sub Command2_Click()
On Error GoTo l
    cdisave.Filter = "(*.bmp)| *.bmp"
    cdisave.ShowSave
    SavePicture FX.Image, cdisave.FileName
l:
End Sub

Private Sub Form_Activate()
txtMessage.SetFocus
End Sub
Private Sub high_Click()
to1.Caption = "100"
FX.Cls
For intCount = to1.Caption To 250
    FX.ForeColor = RGB(intCount + d1.Caption, intCount + d2.Caption, intCount + d3.Caption)
    FX.CurrentX = intCount
    FX.CurrentY = intCount
    FX.Print txtMessage.Text
Next intCount
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Image3.Visible = False
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image4.Visible = False
End Sub

Private Sub Image3_Click()
WindowState = 1
End Sub

Private Sub Image4_Click()
Unload Me
End Sub
Private Sub Image7_Click()
PopupMenu Form2.menu
End Sub
Private Sub List1_Click()
If List1.ListCount <> 0 Then
        FX.Font = List1.Text
    Else
        MsgBox "you have To choose the fonts first"
    End If
End Sub
Private Sub low_Click()
to1.Caption = "230"
FX.Cls
For intCount = to1.Caption To 250
FX.ForeColor = RGB(intCount + d1.Caption, intCount + d2.Caption, intCount + d3.Caption)
    FX.CurrentX = intCount
    FX.CurrentY = intCount
    FX.Print txtMessage.Text
Next intCount
End Sub
Private Sub medium_Click()
to1.Caption = "199"
FX.Cls
For intCount = to1.Caption To 250
    FX.ForeColor = RGB(intCount + d1.Caption, intCount + d2.Caption, intCount + d3.Caption)
    FX.CurrentX = intCount
    FX.CurrentY = intCount
    FX.Print txtMessage.Text
Next intCount
End Sub
Private Sub Picture1_Click()
d1.Caption = "-100"
d2.Caption = "-1"
d3.Caption = "-100"
End Sub
Private Sub Picture10_Click()
d1.Caption = "6300"
d2.Caption = "-1"
d3.Caption = "10"
End Sub

Private Sub Picture11_Click()
d1.Caption = "-100"
d2.Caption = "10"
d3.Caption = "-10"
End Sub

Private Sub Picture12_Click()
d1.Caption = "20"
d2.Caption = "-20"
d3.Caption = "10"
End Sub

Private Sub Picture2_Click()
d1.Caption = "100"
d2.Caption = "-90"
d3.Caption = "-90"
End Sub

Private Sub Picture3_Click()
d1.Caption = "1"
d2.Caption = "1"
d3.Caption = "-100"
End Sub

Private Sub Picture4_Click()
d1.Caption = "-20"
d2.Caption = "20"
d3.Caption = "-30"
End Sub

Private Sub Picture5_Click()
d1.Caption = "-30"
d2.Caption = "-90"
d3.Caption = "-30"
End Sub

Private Sub Picture6_Click()
d1.Caption = "1"
d2.Caption = "20"
d3.Caption = "20"
End Sub

Private Sub Picture7_Click()
d1.Caption = "20"
d2.Caption = "-10"
d3.Caption = "-30"
End Sub

Private Sub Picture8_Click()
d1.Caption = "6300"
d2.Caption = "-1"
d3.Caption = "-100"
End Sub

Private Sub Picture9_Click()
d1.Caption = "-30"
d2.Caption = "-1"
d3.Caption = "-100"
End Sub



Private Sub Timer1_Timer()
If to1.Caption = "230" Then
medium.Enabled = True
low.Enabled = False
high.Enabled = True
ElseIf to1.Caption = "100" Then
medium.Enabled = True
low.Enabled = True
high.Enabled = False
ElseIf to1.Caption = "200" Then
medium.Enabled = False
low.Enabled = True
high.Enabled = True
End If
sim.Caption = FX.FontSize
End Sub
Private Sub txtMessage_Change()
FX.Cls
For intCount = to1.Caption To 250
    FX.ForeColor = RGB(intCount + d1.Caption, intCount + d2.Caption, intCount + d3.Caption)
    FX.CurrentX = intCount
    FX.CurrentY = intCount
    FX.Print txtMessage.Text
Next intCount
End Sub




