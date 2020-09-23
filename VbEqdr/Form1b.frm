VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Equation Draw test"
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Left            =   6150
         Top             =   3450
         Width           =   165
      End
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   150
      TabIndex        =   0
      Text            =   "4*sqr(sqr(5^5)+sqr(5+6^sqr(5^5))-4^3)-3^(4+5)=sqr(x^2+y^2)-4"
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
Dim CX(99) As Integer, CY(99) As Integer

Dim EWidth As Integer, EHeight(99) As Integer

Dim Fonts(99) As Integer

Dim I As Integer

Dim TempStr As String

Dim TempV(99) As String

Dim NxtOpP(99) As Integer

Dim NxtOp(99) As String

Dim PrevOp(99) As String

Dim PrevOpP(99) As Integer

Dim NxtBrP(99) As Integer

Dim SqrH(99) As Integer
Dim SqrIt(99) As Integer
Dim Sqrh1 As Integer

Dim Iter As Integer




Function GetNxtOp(ByVal Astr As String, Optional ByVal Start = 1, Optional ByRef Position) As String
GetNxtOp = ""

For I = Start To Len(Astr)
  TempStr = LCase(Mid(Astr, I, 1))
  
  Select Case TempStr
    Case "+", "=", ":"
      GetNxtOp = TempStr
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
    Case "s"
      If LCase(Mid(Astr, I, 4)) = "sqr(" Then GetNxtOp = "sqr("
      Exit For
  End Select
Next

Position = I
End Function

Function GetPrevOp(ByVal Astr As String, Optional ByVal Start = 1, Optional ByRef Position) As String
GetPrevOp = ""

For I = Start To 1 Step -1
  TempStr = Mid(Astr, I, 1)
  
  Select Case TempStr
    Case "+", "=", ":"
      GetPrevOp = TempStr
      Exit For
    Case "-"
      GetPrevOp = "-"
      Exit For
    Case "/"
      GetPrevOp = "/"
      Exit For
    Case "*"
      GetPrevOp = "*"
      Exit For
    Case "^"
      GetPrevOp = "^"
      Exit For
    Case "(", ")"
      GetPrevOp = TempStr
      Exit For
    Case "s"
      If LCase(Mid(Astr, I, 4)) = "sqr(" Then GetPrevOp = "sqr("
      Exit For
  End Select
Next

Position = I + 1
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

ErrNonZ = TempBr

FindBr = I
End Function


Function Parse(ByVal Astr As String, Optional Start As Integer = 1, Optional defFonts As Integer = 10) As Boolean
Parse = True



Dim C As Integer

C = Start

Do Until C > Len(Astr)

  NxtOp(Iter) = GetNxtOp(Astr, C, NxtOpP(Iter))
  PrevOp(Iter) = GetPrevOp(Astr, NxtOpP(Iter) - 1, PrevOpP(Iter))
    
  Select Case PrevOp(Iter)
    Case ""
      TempV(Iter) = Mid(Astr, PrevOpP(Iter), NxtOpP(Iter) - PrevOpP(Iter))
      DrawNumber TempV(Iter), CX(Iter), CY(Iter)
    Case "+", "-", "=", ":"
      TempV(Iter) = Mid(Astr, PrevOpP(Iter), NxtOpP(Iter) - PrevOpP(Iter))
      DrawNumber PrevOp(Iter) & TempV(Iter), CX(Iter), CY(Iter)
    
    Case "*"
      TempV(Iter) = Mid(Astr, PrevOpP(Iter), NxtOpP(Iter) - PrevOpP(Iter))
      DrawNumber Chr(183) & TempV(Iter), CX(Iter), CY(Iter)
      
    Case "sqr("
        PrevOpP(Iter) = PrevOpP(Iter) + 3
        If NxtOp(Iter) = "(" Then
          NxtOpP(Iter) = FindBr(Astr, NxtOpP(Iter)) + 1
          NxtOp(Iter) = GetNxtOp(Astr, NxtOpP(Iter))
        End If
        
      TempV(Iter) = Mid(Astr, PrevOpP(Iter), NxtOpP(Iter) - PrevOpP(Iter) - 1)
        
      Fonts(Iter + 1) = Fonts(Iter)
      CX(Iter + 1) = CX(Iter) + 12 * (Fonts(Iter) / 18)
      CY(Iter + 1) = CY(Iter)
      TempV(Iter + 1) = TempV(Iter)
      EHeight(Iter + 1) = EHeight(Iter) + 6 * (Fonts(Iter) / 18)
      
      Sqrh1 = Sqrh1 + 1
      SqrIt(Sqrh1) = EHeight(Iter)
            
      Iter = Iter + 1
      Parse TempV(Iter), 1
      Iter = Iter - 1

      For I = Sqrh1 To 1 Step -1
        If SqrIt(I) > SqrIt(Sqrh1 - 1) Then SqrIt(I - 1) = SqrIt(I)
      Next
      
      SqrH(Iter) = SqrIt(Sqrh1) - EHeight(Iter)
      Sqrh1 = Sqrh1 - 1

      P1.Line (CX(Iter) + 3 * (Fonts(Iter) / 18), CY(Iter) + 15 * (Fonts(Iter) / 18))-(CX(Iter) + 6 * (Fonts(Iter) / 18), CY(Iter) + 25 * (Fonts(Iter) / 18))
      P1.Line (CX(Iter) + 6 * (Fonts(Iter) / 18), CY(Iter) + 25 * (Fonts(Iter) / 18))-(CX(Iter) + 12 * (Fonts(Iter) / 18), CY(Iter) - SqrH(Iter))
      P1.Line (CX(Iter) + 12 * (Fonts(Iter) / 18), CY(Iter) - SqrH(Iter))-(CX(Iter + 1), CY(Iter) - SqrH(Iter))
      P1.Line (CX(Iter + 1), CY(Iter) - SqrH(Iter))-(CX(Iter + 1) + 3 * (Fonts(Iter) / 18), CY(Iter) - SqrH(Iter) + 9 * (Fonts(Iter) / 18))
            
      CX(Iter) = CX(Iter + 1) + 6 * (Fonts(Iter) / 18)
      CY(Iter) = CY(Iter + 1)
      
    Case "^"
    
      Do
        If NxtOp(Iter) = "(" Then
          NxtOpP(Iter) = FindBr(Astr, NxtOpP(Iter)) + 1
          NxtOp(Iter) = GetNxtOp(Astr, NxtOpP(Iter))
        End If
        
        If NxtOp(Iter) = "sqr(" Then
          NxtOpP(Iter) = FindBr(Astr, NxtOpP(Iter)) + 1
          NxtOp(Iter) = GetNxtOp(Astr, NxtOpP(Iter))
        End If
      
        If NxtOp(Iter) = "^" Then
          Do
            NxtOp(Iter) = GetNxtOp(Astr, NxtOpP(Iter) + 1, NxtOpP(Iter))
          Loop Until NxtOp(Iter) <> "^"
        End If
      Loop Until NxtOp(Iter) <> "(" And NxtOp(Iter) <> "sqr("
      
      TempV(Iter) = Mid(Astr, PrevOpP(Iter), NxtOpP(Iter) - PrevOpP(Iter))
      If Fonts(Iter) > 9 Then Fonts(Iter + 1) = Fonts(Iter) - 3 Else: Fonts(Iter + 1) = Fonts(Iter)
      CX(Iter + 1) = CX(Iter)
      CY(Iter + 1) = CY(Iter) - 9
      EHeight(Iter + 1) = EHeight(Iter) + 9
      TempV(Iter + 1) = TempV(Iter)
      Iter = Iter + 1
      Parse TempV(Iter), 1
      Iter = Iter - 1
      CX(Iter) = CX(Iter + 1)
      CY(Iter) = CY(Iter + 1) + 9
      
    Case ")"
      GoTo 10
      
    Case "("
      DrawNumber "(", CX(Iter), CY(Iter)
      NxtBrP(Iter) = FindBr(Astr, C - 1)
      TempV(Iter) = Mid(Astr, PrevOpP(Iter), NxtBrP(Iter) - PrevOpP(Iter))
      Fonts(Iter + 1) = Fonts(Iter)
      CX(Iter + 1) = CX(Iter)
      CY(Iter + 1) = CY(Iter)
      EHeight(Iter + 1) = EHeight(Iter)
      TempV(Iter + 1) = TempV(Iter)
      Iter = Iter + 1
      Parse TempV(Iter), 1
      Iter = Iter - 1
      CX(Iter) = CX(Iter + 1)
      CY(Iter) = CY(Iter + 1)
      DrawNumber ")", CX(Iter), CY(Iter)
      C = NxtBrP(Iter): GoTo 11
      
  End Select
10:
  C = NxtOpP(Iter) + 1
11:
  If EHeight(Iter) > SqrIt(Sqrh1) Then SqrIt(Sqrh1) = EHeight(Iter)
Loop
End Function

Function DrawNumber(ByVal Astr As String, Xpos As Integer, Ypos As Integer)
  P1.CurrentX = Xpos
  P1.CurrentY = Ypos
  P1.FontSize = Fonts(Iter)
  If Fonts(Iter) < 13 Then P1.FontBold = True Else: P1.FontBold = False
  P1.Print Astr
  Xpos = Xpos + P1.TextWidth(Astr)
  EWidth = EWidth + P1.TextWidth(Astr)
End Function

Private Sub Form_Load()
PrevOpP(0) = 1
'Shape1.Top = 10
'Shape1.Left = 10
End Sub

Private Sub Text1_Change()
P1.Cls
CX(0) = 10
CY(0) = P1.Height / 2
Fonts(0) = 18
P1.FontSize = Fonts(0)
EWidth = 0

For I = 0 To UBound(EHeight)
  EHeight(I) = 0
Next

EHeight(0) = P1.TextHeight("A")
Iter = 0
Sqrh1 = 0

For I = 0 To UBound(SqrIt)
  SqrIt(I) = 0
Next

Parse Text1
'For I = 99 To 0 Step -1
  'If EHeight(I) <> 0 Then
    'Shape1.Height = EHeight(I)
    'Shape1.Top = 10 - (Shape1.Height - P1.TextHeight("A"))
    'Label1.Caption = EHeight(I)
    'Exit For
  'End If
'Next
End Sub
