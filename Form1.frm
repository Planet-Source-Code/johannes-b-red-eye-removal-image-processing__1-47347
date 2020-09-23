VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Red-eye removal 1.0 by Johannes B 2003"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox TempPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4920
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   20
      Top             =   3960
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Select eye..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Top             =   3120
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "D"
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      ToolTipText     =   "Use default value"
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "D"
      Height          =   255
      Left            =   7440
      TabIndex        =   8
      ToolTipText     =   "Use default value"
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "D"
      Height          =   255
      Left            =   7440
      TabIndex        =   9
      ToolTipText     =   "Use default value"
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Caption         =   "C"
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      ToolTipText     =   "Center scroll"
      Top             =   4440
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Undo"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   2640
      Width           =   1695
   End
   Begin VB.HScrollBar hred 
      Height          =   255
      LargeChange     =   25
      Left            =   4320
      Max             =   0
      Min             =   255
      TabIndex        =   15
      Top             =   600
      Value           =   60
      Width           =   3135
   End
   Begin VB.HScrollBar hother 
      Height          =   255
      LargeChange     =   25
      Left            =   4320
      Max             =   255
      TabIndex        =   14
      Top             =   1200
      Value           =   160
      Width           =   3135
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto preview"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.HScrollBar hbright 
      Height          =   255
      LargeChange     =   25
      Left            =   4320
      Max             =   100
      Min             =   -100
      TabIndex        =   10
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Process from red"
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      ToolTipText     =   "If enabled the red channel will be included while processing the output colors"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Restore image"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   3855
   End
   Begin VB.PictureBox PC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1365
         Left            =   360
         OLEDropMode     =   1  'Manual
         Picture         =   "Form1.frx":08CA
         ScaleHeight     =   91
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   191
         TabIndex        =   4
         Top             =   1200
         Width           =   2865
         Begin VB.Shape Shape3 
            BorderStyle     =   3  'Dot
            DrawMode        =   16  'Merge Pen
            Height          =   255
            Left            =   480
            Top             =   600
            Width           =   255
            Visible         =   0   'False
         End
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   3615
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4335
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   3495
      Left            =   4200
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Red tolerance"
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Tolerance for other colors"
      Height          =   255
      Left            =   4320
      TabIndex        =   17
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Brightnes out"
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   1560
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pixel

Dim Rred
Dim Ggreen
Dim Bblue

Dim RR1
Dim GG1
Dim BB1


Dim Q As Integer
Dim Q2 As Integer
Dim Q3 As Integer
Dim Q4 As Integer


Dim XXX As Integer
Dim YYY As Integer


Dim CurX
Dim CurY

Dim JB As Byte

Dim RER As Byte
Dim RX As Integer
Dim RY As Integer

Dim One As Integer
Dim Two As Integer
Private Sub GetRGB(ByVal Col As String)
On Error Resume Next
    Bblue = Col \ (256 ^ 2)
    Ggreen = (Col - Bblue * 256 ^ 2) \ 256
    Rred = (Col - Bblue * 256 ^ 2 - Ggreen * 256)
End Sub
Sub RedEyeRemoval(RRX, RRY, RRXX, RRYY)
On Error Resume Next
'Convert hscroll values to integers for faster processing
Q = hred.Value
Q2 = hother.Value
Q3 = hbright.Value
Q4 = Check2.Value

Screen.MousePointer = 11
For YYY = RRY To RRYY - 1
For XXX = RRX To RRXX - 1
'Read pixel
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
'Separate RGB channels
GetRGB Pixel

'Is this a red eye?
If Val(Rred) > Q And Val(Ggreen) < Q2 And Val(Bblue) < Q2 Then
'Process the new color to replace the red eye with
If Q4 = 0 Then
Rred = (Val(Ggreen) + Val(Bblue)) / 2
Else
Rred = (Val(Rred) + Val(Ggreen) + Val(Bblue)) / 3
End If
'Brightnes
Rred = Val(Rred) + Q3
'Lower/upper limit
If Rred < 0 Then Rred = 0
If Rred > 255 Then Rred = 255
'Draw new pixel
SetPixelV Picture1.HDC, XXX, YYY, RGB(Rred, Rred, Rred)
End If


Next
Picture1.Refresh
Next
Picture1.Refresh
Screen.MousePointer = 0
RER = 0
End Sub
Sub CopyFromSecond()
Picture1.Width = TempPic.ScaleWidth
Picture1.Height = TempPic.ScaleHeight
BitBlt Picture1.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, TempPic.HDC, 0, 0, vbSrcCopy
Picture1.Refresh
End Sub
Sub CopyToSecond()
TempPic.Width = Picture1.ScaleWidth
TempPic.Height = Picture1.ScaleHeight
BitBlt TempPic.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, vbSrcCopy
TempPic.Refresh
End Sub
Sub UppdateScroll()
If Picture1.Width <= PC.ScaleWidth Then
HScroll1.Enabled = False
Else
HScroll1.Enabled = True
End If

If Picture1.Height <= PC.ScaleHeight Then
VScroll1.Enabled = False
Else
VScroll1.Enabled = True
End If

VScroll1.Max = Picture1.ScaleHeight - PC.ScaleHeight
HScroll1.Max = Picture1.ScaleWidth - PC.ScaleWidth

If HScroll1.Enabled = False Then
Picture1.Left = (PC.Width / 2) - (Picture1.Width / 2)
End If
If VScroll1.Enabled = False Then
Picture1.Top = (PC.Height / 2) - (Picture1.Height / 2)
End If

HScroll1.Value = 0
VScroll1.Value = 0
End Sub



Private Sub Check1_Click()
If Check1.Value = 1 Then
Command8.Value = True
End If
End Sub

Private Sub Check2_Click()
If Check1.Value = 1 Then Command8.Value = True
End Sub

Private Sub Command1_Click()
hred.Value = "60"
End Sub

Private Sub Command3_Click()
Form1.CopyFromSecond
End Sub

Private Sub Command35_Click()
If HScroll1.Enabled = True Then HScroll1.Value = HScroll1.Max / 2
If VScroll1.Enabled = True Then VScroll1.Value = VScroll1.Max / 2
End Sub

Private Sub Command4_Click()
If RER = 0 Then
If Two = 0 Then MsgBox "Now move mouse pointer to an eye, right button + drag = resize selection, left button = select area!", vbInformation
Two = 1
RER = 1
Shape3.Visible = True
Command4.Caption = "Cancel"

Else
Shape3.Visible = False
Command4.Caption = "Select eye..."
RER = 0
End If
End Sub


Private Sub Command5_Click()
Picture1.Cls
End Sub

Private Sub Command6_Click()
hother.Value = "160"
End Sub

Private Sub Command7_Click()
hbright.Value = "0"
End Sub

Private Sub Command8_Click()
If Shape3.Visible = False Then
MsgBox "No eye area selected! Click on Select eye to make a selection", vbInformation
Exit Sub
End If
Form1.CopyFromSecond
Form1.RedEyeRemoval Form1.Shape3.Left, Form1.Shape3.Top, Form1.Shape3.Left + Form1.Shape3.Width, Form1.Shape3.Top + Form1.Shape3.Height

End Sub

Private Sub Form_Load()
UppdateScroll
End Sub






Private Sub hbright_Change()
If Check1.Value = 1 Then Command8.Value = True
End Sub

Private Sub hother_Change()
If Check1.Value = 1 Then Command8.Value = True
End Sub

Private Sub hred_Change()
If Check1.Value = 1 Then Command8.Value = True
End Sub


Private Sub HScroll1_Change()
Picture1.Left = 0 - HScroll1.Value
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
JB = 1
CurX = X
CurY = Y
If RER = 1 And Button = vbRightButton Then
RX = X
RY = Y
End If
If RER = 1 And Button = vbLeftButton Then
Command4.Caption = "Select eye..."
RER = 0
CopyToSecond
End If
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If JB = 1 And RER = 0 Then
If HScroll1.Enabled = True Then HScroll1.Value = HScroll1.Value + CurX - X
If VScroll1.Enabled = True Then VScroll1.Value = VScroll1.Value + CurY - Y
End If
If RER = 1 Then
If Button = vbRightButton Then
Shape3.Width = Shape3.Left + (X - (Shape3.Left * 2))
Shape3.Height = Shape3.Top + (Y - (Shape3.Top * 2))
Else
Shape3.Left = X
Shape3.Top = Y
End If
End If
End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
JB = 0
End Sub


Private Sub VScroll1_Change()
Picture1.Top = 0 - VScroll1.Value
End Sub


