VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binary & Decimal Converter"
   ClientHeight    =   1320
   ClientLeft      =   5115
   ClientTop       =   4740
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "# ##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   255
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "This is the conversion of the number that you entered above"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Decimal"
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   "Click this if you want to convert binary to decimal"
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Binary"
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Click this if you want to convert decimal to binary"
      Top             =   840
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "# ##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      HideSelection   =   0   'False
      Left            =   120
      MaxLength       =   9
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "Enter the number that you want to convert here"
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Number To Convert:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Convert to:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Conversion:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the best way to convert numbers from binary - decimal and back
'Theese first two proceedures are the basic conversion functions
'They can convert numbers in base 2,3,4,5,6,7,8,9 and 10!

Public Property Get Deciml(InptB As Variant, BaseB As Variant) As Variant

B = InptB
E = 0
F = 0


Do
A = Right(B, 1) 'this is where the conversion happens
B = Left(B, Len(B) - 1) 'this is where the conversion happens

C = BaseB ^ F 'conversion
D = A * C 'conversion
E = E + D 'conversion

F = F + 1 'counter
Loop Until B = ""

Deciml = E
End Property



Public Property Get Binary(InptD As Variant, BaseD As Variant) As Variant
Binary = ""
G = InptD
Do
Binary = (G Mod BaseD) & Binary 'conversion
G = G \ BaseD 'conversion
Loop Until G = 0
End Property

Sub AutoSelect(SelObject As Control)
    On Error Resume Next
    SelObject.SelStart = 0
    If TypeOf SelObject Is TextBox Then
        SelObject.SelLength = Len(SelObject.Text)
    End If
End Sub

Private Sub Option1_Click()
Text1.MaxLength = 9
Text1.Text = Left(Text1.Text, Text1.MaxLength)
Text1_Change
Text1.SetFocus
End Sub

Private Sub Option2_Click()
Text1.MaxLength = 49
Text1_Change


IlglChr2 = InStr(Text1.Text, "2")
IlglChr3 = InStr(Text1.Text, "3")
IlglChr4 = InStr(Text1.Text, "4")
IlglChr5 = InStr(Text1.Text, "5")
IlglChr6 = InStr(Text1.Text, "6")
IlglChr7 = InStr(Text1.Text, "7")
IlglChr8 = InStr(Text1.Text, "8")
IlglChr9 = InStr(Text1.Text, "9")

If IlglChr2 Or IlglChr3 Or IlglChr4 Or IlglChr5 Or IlglChr6 Or IlglChr7 Or IlglChr8 Or IlglChr9 Then
Response = MsgBox("Sorry, you have an illegal character in the 'Number To Convert' space.  It will be  reset to zero." & vbNewLine & "Choose Cancel to abort.", vbOKCancel, Form1.Caption & ":")

If Response = vbCancel Then
Option1.Value = True
Else
Text1.Text = "0"
End If

End If
Text1.SetFocus

End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then

Select Case Option1.Value

Case True
Text2.Text = Binary(Text1.Text, 2)
Case False
Text2.Text = Deciml(Text1.Text, 2)
End Select
End If

End Sub

Private Sub Text1_GotFocus()
AutoSelect Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case Option1.Value

Case True 'convert to binary
    Select Case KeyAscii
        Case 8 'backspace
        Case 48 To 57 '0 and 1
        Case Else
            KeyAscii = 0
    End Select
Case False 'convert to decimal
    Select Case KeyAscii
        Case 8 'backspace
        Case 48 To 49 '0 to 9
        Case Else
            KeyAscii = 0
    End Select

End Select

End Sub

Private Sub Text2_GotFocus()
AutoSelect Text2
End Sub
