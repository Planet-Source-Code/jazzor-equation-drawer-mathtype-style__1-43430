VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   518
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   4140
      Left            =   75
      ScaleHeight     =   272
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   1
      Top             =   750
      Width           =   6615
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   3600
         Top             =   1575
         Width           =   165
      End
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   150
      TabIndex        =   0
      Text            =   "45^(3+64)"
      Top             =   150
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   240
      Left            =   6525
      TabIndex        =   2
      Top             =   300
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CX As Integer, CY As Integer

Dim EWidth As Integer, EHeight As Integer

Dim Fonts As Integer

Dim I As Integer

Dim TempStr As String

Dim TempV As String

Dim NxtOpP As Integer

Dim NxtOp As String

Dim PrevOp As String

Dim NxtBrP As Integer



Function GetNxtOp(ByVal Astr As String, Optional ByVal Start = 1, Optional ByRef Position) As String
For I = Start To Len(Astr)
  TempStr = Mid(Astr, I, 1)
  
  Select Case TempStr
    Case "+"
      GetNxtOp = "+"
      Exit For
    Case "-"
      GetNxtOp = "-"
      Exit For
    Case "/"
      GetNxtOp = "/"
      Exit For
    Case "*"
      GetNxtOp = "*"
      Exit For
    Case "^"
      GetNxtOp = "^"
      Exit For
    Case "("
      GetNxtOp = "("
      Exit For
  End Select
Next

Position = I
End Function

Function FindBr(ByVal Astr As String, Optional ByVal Start = 1, Optional ByRef ErrNonZ = 0) As Integer
Dim TempBr As Integer
TempBr = 0

For I = Start To Len(Astr)
  TempStr = Mid(Astr, I, 1)
  If TempStr = "(" Then TempBr = TempBr + 1
  If TempStr = ")" Then TempBr = TempBr - 1
  If TempBr = 0 Then Exit For
Next

ErrNonZ = br

FindBr = I
End Function


Function Parse(ByVal Astr As String, Optional Start As Integer = 1, Optional defFonts As Integer = 10) As Boolean
Parse = True



Dim C As Integer

C = Start

TempV = ""

Do Until C > Len(Astr)

  NxtOp = GetNxtOp(Astr, C, NxtOpP)
  TempV = ""
  For I = C To NxtOpP - 1
    TempStr = Mid(Astr, I, 1)
    TempV = TempV & TempStr
  Next
    
  Select Case PrevOp
    Case ""
      DrawNumber TempV, CX, CY
      PrevOp = NxtOp
      C = NxtOpP + 1
    Case "+"
      DrawNumber "+" & TempV, CX, CY
      PrevOp = NxtOp
      C = NxtOpP + 1
    'Case "("
      'NxtBrP = FindBr(Astr, C)
      'TempV = Mid(Astr, C, NxtBrP - C)
      'Form1.Caption = TempV
    Case "^"
      Select Case NxtOp
        Case "("
          NxtBrP = FindBr(Astr, C)
          TempV = Mid(Astr, C + 1, NxtBrP - C - 1)
          PrevOp = "^"
          Parse TempV, 1
          C = NxtBrP + 1
        Case Else
          DrawPower TempV, CX, CY
          PrevOp = NxtOp
          C = NxtOpP + 1
      End Select
  End Select
  
Loop


'C = NxtOpP


'Select Case NxtOp
  'Case "^"
    'NxtOp = GetNxtOp(Astr, NxtOpP + 1, NxtOpP)
    'For I = C To NxtOpP - 1
      'TempStr = Mid(Astr, I, 1)
      'DrawPower TempStr, CX, CY
    'Next
'End Select
'Shape1.Height = EHeight
'Shape1.Left = CX
'Shape1.Top = CY
Label1.Caption = EHeight
End Function

Function DrawNumber(ByVal Astr As String, Xpos As Integer, Ypos As Integer)
  P1.CurrentX = Xpos
  P1.CurrentY = Ypos
  P1.FontSize = Fonts
  P1.Print Astr
  Xpos = Xpos + P1.TextWidth(Astr)
  EWidth = EWidth + P1.TextWidth(Astr)
End Function

Function DrawPower(Astr As String, Xpos As Integer, Ypos As Integer)
  P1.CurrentX = Xpos
  P1.CurrentY = Ypos
  P1.FontSize = Fonts - 2
  P1.Print Astr
  Xpos = Xpos + P1.TextWidth(Astr)
  EWidth = EWidth + P1.TextWidth(Astr)
  EHeight = EHeight + 2
End Function

Private Sub Form_Load()
CX = 10
CY = 10
EWidth = 0
EHeight = P1.TextHeight("A")
End Sub

Private Sub Text1_Change()
P1.Cls
P1.CurrentX = 0
P1.CurrentY = 0
CX = 10
CY = 10
Fonts = 10
EWidth = 0
EHeight = P1.TextHeight("A")
Parse Text1
End Sub
