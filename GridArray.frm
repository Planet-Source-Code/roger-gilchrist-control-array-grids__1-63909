VERSION 5.00
Begin VB.Form GridArray 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optDemo 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Steps for QBColor"
      Top             =   2880
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdDemo 
      Caption         =   "cmdDemo"
      Height          =   375
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Apply QBColor"
      Top             =   360
      Width           =   375
   End
   Begin VB.Shape shpOrange 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "GridArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lStep     As Long



Private Sub cmdDemo_Click(Index As Integer)

  Dim I As Long
LockWindowUpdate Me.hWnd
  For I = 0 To shpOrange.UBound Step lStep
    shpOrange(I).BackColor = QBColor(Index Mod 16)
    shpOrange(I).Refresh
  Next I
LockWindowUpdate 0
End Sub

Private Sub Form_Load()

  Dim I As Long

  lStep = 1
  GenerateGrid optDemo, 1, 10
  GenerateGrid cmdDemo, 4, 4, cmdDemo(0).Width / 6
'Using Ubound allows you to manipulate the controls created
  For I = 0 To cmdDemo.UBound
    cmdDemo(I).Caption = I
    cmdDemo(I).BackColor = QBColor(I)
  Next I
  For I = 0 To optDemo.UBound
    optDemo(I).Caption = "Step " & I + 1
  Next I
  optDemo(0).Value = True
'forces the first Option to True;
'Load has an interesting effect on OptionButton.Value property
'if the root control is Value = True then the 1st 'Load'ed control takes
'the Value and because of the way OptionButtons are linked together the root control is set to False
'from then on the new Loaded controls take .Value = False from the root control.

End Sub

Private Sub Form_Resize()

  shpOrange(0).BackColor = 33023
  GenerateGrid shpOrange, Me.Width \ shpOrange(0).Width + 1, Me.Height \ shpOrange(0).Height
  Caption = shpOrange.UBound & " dots"

End Sub

Private Sub optDemo_Click(Index As Integer)

  lStep = Index + 1

End Sub

':)Code Fixer V4.0.29 (Wednesday, 04 January 2006 11:20:29) 2 + 49 = 51 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 133302322223333232|33332222222222222222222222222202|1112222|2221222|222222222233|111111111111|1122222222220|333333|
