VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm3DText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D-Text: Programmed By Kenneth Roberts (AKA Bugz)"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdColorPicker 
      Left            =   3360
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtToDisplay 
      Height          =   285
      Left            =   3840
      TabIndex        =   8
      Top             =   1560
      Width           =   3015
   End
   Begin VB.OptionButton optAlign 
      Caption         =   "Right-Alignment"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.OptionButton optAlign 
      Caption         =   "Center-Alignment"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton optAlign 
      Caption         =   "Left-Alignment"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Back Color:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Front Color:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00130095&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblColor 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   12
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frm3DText.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Shape shpHUDCenter 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image cmdArrow 
      Height          =   510
      Index           =   3
      Left            =   5680
      Picture         =   "frm3DText.frx":020F
      Top             =   3180
      Width           =   300
   End
   Begin VB.Image cmdArrow 
      Height          =   510
      Index           =   2
      Left            =   5680
      Picture         =   "frm3DText.frx":0A49
      Top             =   2320
      Width           =   300
   End
   Begin VB.Image cmdArrow 
      Height          =   300
      Index           =   1
      Left            =   5160
      Picture         =   "frm3DText.frx":1283
      Top             =   2880
      Width           =   510
   End
   Begin VB.Image cmdArrow 
      Height          =   300
      Index           =   0
      Left            =   6000
      Picture         =   "frm3DText.frx":1AE5
      Top             =   2880
      Width           =   510
   End
   Begin VB.Line Lines 
      Index           =   4
      X1              =   4920
      X2              =   6720
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape shpHudBG 
      Height          =   1935
      Left            =   4920
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "3D Text To Display:"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Shape shpHUD 
      Height          =   1455
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Alignment:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblFront 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.Shape shpBG 
      Height          =   1215
      Left            =   2040
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label lblBack 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2220
      TabIndex        =   2
      Top             =   350
      Width           =   4575
   End
   Begin VB.Label lblBackground 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2060
      TabIndex        =   0
      Top             =   230
      Width           =   4815
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "HUD:"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   9
      Top             =   1950
      Width           =   1935
   End
   Begin VB.Label lbl 
      Caption         =   "3-D Text Display Window:"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   10
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frm3DText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'3-D Text coded by Kenneth Roberts (AKA Bugz)
'Date Programmed: December 05, 2006
'Date Released: December 05, 2006

Private Sub cmdArrow_Click(Index As Integer)
'For more information on how the index is used, please view the
'optAlign_Click Sub
'-
'We then Align both of the labels, the front label (lblFront) and the
'back label (lblBack).
'What gives this the 3-D look, is that one label (lblBack)
'is placed slightly higher/lower and a little to the left/right of lblFront.

Select Case Index
  Case 0 'Right Arrow
    lblBack.Left = lblBack.Left + 10
  Case 1 'Left Arrow
    lblBack.Left = lblBack.Left - 10
  Case 2 'Up Arrow
    lblBack.Top = lblBack.Top - 10
  Case 3 'Down Arrow
    lblBack.Top = lblBack.Top + 10
End Select
End Sub

Private Sub Form_Load()
lblBackground.BackColor = vbBlack
lblFront.ForeColor = lblColor(0).BackColor '&HFF& OR vbRed
lblBack.ForeColor = lblColor(1).BackColor '&H130095
txtToDisplay.Text = "3-D Text"
End Sub

Private Sub lblColor1_Click()
End Sub

Private Sub lblColor_Click(Index As Integer)
On Error GoTo ErrColor: 'Add a little Error Routine.
'Show our color picker with Microsoft's nifty Common Dialog control.
cdColorPicker.ShowColor

'Heres that index again.. Darn thing wont go away for anything.
Select Case Index
  Case 0
    lblColor(0).BackColor = cdColorPicker.Color
    lblFront.ForeColor = lblColor(0).BackColor
  Case 1
    lblColor(1).BackColor = cdColorPicker.Color
    lblFront.ForeColor = lblColor(1).BackColor
End Select

Exit Sub
ErrColor:
'We usually come here, sometimes when the cancel button is hit.
'Or something other of the sort.
MsgBox "There was an error choosing your color.", vbOKOnly + vbCritical, "Error"
End Sub

Private Sub optAlign_Click(Index As Integer)
'The optAlign controls are for the 3-D Text Alignment.
'We base our option off of the index provided.
'If we click on the first Option (its index is 0).
'Second Option = index 1, etc
Select Case Index
  Case 0 'Left Alignment
    lblFront.Alignment = 0
    lblBack.Alignment = 0
  Case 1 'Center Alignment
    lblFront.Alignment = 2
    lblBack.Alignment = 2
  Case 2 'Right Alignment
    lblFront.Alignment = 1
    lblBack.Alignment = 1
End Select
End Sub

Private Sub txtToDisplay_Change()
'Set the front and back label (lblFront) & (lblBack) caption as the same
'text of the txtToDisplay textbox.
lblFront.Caption = txtToDisplay.Text
lblBack.Caption = txtToDisplay.Text
End Sub
