VERSION 5.00
Begin VB.Form calci 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5445
   DrawMode        =   16  'Merge Pen
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command16 
      Caption         =   "Back"
      Height          =   615
      Left            =   360
      TabIndex        =   27
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command15 
      Caption         =   "View last"
      Height          =   615
      Left            =   360
      TabIndex        =   25
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   360
      TabIndex        =   23
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0FFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   22
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   21
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   20
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   19
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   18
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   17
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   15
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   14
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   13
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   12
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   10
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   360
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   9
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   2280
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   1320
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   7
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   360
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   2280
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1320
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   360
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2280
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1320
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   360
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   2895
      Left            =   360
      TabIndex        =   26
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sourav Agarwal && Sunny SIngh"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   2280
      TabIndex        =   24
      Top             =   4680
      Width           =   2655
   End
End
Attribute VB_Name = "calci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op1, op2 As Double
Dim d, cnt As Integer
Dim opr, lst As String

Private Sub Command1_Click(Index As Integer)
If d = 1 Then
Text1.Text = Command1(Index).Caption
d = 0
Else
Text1.Text = Text1.Text + Command1(Index).Caption
End If
lst = lst + Command1(Index).Caption
End Sub

Private Sub Command10_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
Exit Sub
Else
If Text1.Text <> 0 Then
Text1.Text = 1 / Val(Text1.Text)
Else
Text1.Text = "INFINITY"
Exit Sub
End If
End If
lst = lst + "<-reciproical is-> " + Text1.Text + "<finish> "
End Sub

Private Sub Command11_Click()
If cnt = 0 And Text1.Text = "" Then
Exit Sub
Else
op2 = Val(Text1.Text)
Select Case opr
Case "+":
If Text1.Text = "" Then
opr2 = 0
End If
Text1.Text = op1 + op2
Case "*":
If (Text1.Text = "") Then
op2 = 1
End If
Text1.Text = op1 * op2
Case "-":
If Text1.Text = "" Then
opr2 = 0
End If
Text1.Text = op1 - op2
Case "/":
If op2 = 0 Then
Text1.Text = "INFINTY"
Exit Sub
Else
If Text1.Text = "" Then op2 = 1
End If
Text1.Text = op1 / op2
End Select
End If
d = 1
op1 = 0
op2 = 0
opr = ""
cnt = 0
lst = lst + " =" + Text1.Text + "<finish>   "
End Sub

Private Sub Command12_Click()
If Text1.Visible = True Then
If Len(Text1.Text) > 0 Then
Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
Else: Exit Sub
End If
ElseIf Text2.Visible = True Then
If Len(Text2.Text) > 0 Then
Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
Else: Exit Sub
End If
End If
lst = lst + " <-"
End Sub

Private Sub Command13_Click()
Text1.Text = ""
op1 = 0
op2 = 0
opr = ""
cnt = 0
lst = ""
Text1.SetFocus
End Sub

Private Sub Command14_Click()
Text1.Text = ""
op1 = 0
op2 = 0
opr = ""
cnt = 0
Text1.SetFocus
End Sub

Private Sub Command15_Click()
Label2.Visible = True
Command1(0).Visible = False
Command1(1).Visible = False
Command1(2).Visible = False
Command1(3).Visible = False
Command1(4).Visible = False
Command1(5).Visible = False
Command1(6).Visible = False
Command1(7).Visible = False
Command1(8).Visible = False
Command1(9).Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
Command11.Visible = False
Label2.Caption = lst
Command15.Visible = False
Command16.Visible = True
End Sub

Private Sub Command16_Click()
Command1(0).Visible = True
Command1(1).Visible = True
Command1(2).Visible = True
Command1(3).Visible = True
Command1(4).Visible = True
Command1(5).Visible = True
Command1(6).Visible = True
Command1(7).Visible = True
Command1(8).Visible = True
Command1(9).Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command9.Visible = True
Command10.Visible = True
Command11.Visible = True
Command15.Visible = True
Command16.Visible = False
Label2.Visible = False
End Sub

Private Sub Command2_Click()
If Val(Text1.Text) > 0 Then
Text1.Text = "-" & Text1.Text
ElseIf Val(Text1.Text) < 0 Then
Text1.Text = Right(Text1.Text, Len(Text1.Text) - 1)
ElseIf (Text1.Text = " ") Then
Exit Sub
End If
lst = lst + " (previous sign changed)"
End Sub

Private Sub Command3_Click()
If InStr(Text1.Text, ".") Then
Exit Sub
Else
Text1.Text = Text1.Text + "."
End If
lst = lst + "."
End Sub

Private Sub Command4_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
Exit Sub
Else
If cnt = 0 Then
op1 = Val(Text1.Text)
Else
If opr = "+" Then
op1 = op1 + Val(Text1.Text)
ElseIf opr = "-" Then
op1 = op1 - Val(Text1.Text)
ElseIf opr = "*" Then
op1 = op1 * Val(Text1.Text)
ElseIf opr = "/" Then
If Text1.Text = 0 Then
Text1.Text = "infinity"
Exit Sub
Else
op1 = op1 / Val(Text1.Text)
End If
End If
End If
opr = "/"
Text1.Text = ""
cnt = cnt + 1
End If
lst = lst + " /"
End Sub

Private Sub Command5_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
Exit Sub
Else
If cnt = 0 Then
op1 = Val(Text1.Text)
Else
If opr = "+" Then
op1 = op1 + Val(Text1.Text)
ElseIf opr = "-" Then
op1 = op1 - Val(Text1.Text)
ElseIf opr = "*" Then
op1 = op1 * Val(Text1.Text)
ElseIf opr = "/" Then
If Text1.Text = 0 Then
Text1.Text = "infinity"
Exit Sub
Else
op1 = op1 / Val(Text1.Text)
End If
End If
End If
opr = "*"
Text1.Text = ""
cnt = cnt + 1
End If
lst = lst + " *"
End Sub

Private Sub Command6_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
Exit Sub
Else
If cnt = 0 Then
op1 = Text1.Text
Else
If opr = "+" Then
op1 = op1 + Val(Text1.Text)
ElseIf opr = "-" Then
op1 = op1 - Val(Text1.Text)
ElseIf opr = "*" Then
op1 = op1 * Val(Text1.Text)
ElseIf opr = "/" Then
If Text1.Text = 0 Then
Text1.Text = "infinity"
Exit Sub
Else
opr1 = opr1 / Val(Text1.Text)
End If
End If
End If
Text1.Text = ""
opr = "-"
cnt = cnt + 1
End If
lst = lst + " -"
End Sub

Private Sub Command7_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
Exit Sub
Else
If cnt = 0 Then
op1 = Text1.Text
Else
If opr = "+" Then
op1 = op1 + Val(Text1.Text)
ElseIf opr = "-" Then
op1 = op1 - Val(Text1.Text)
ElseIf opr = "*" Then
op1 = op1 * Val(Text1.Text)
ElseIf opr = "/" Then
If Text1.Text = 0 Then
Text1.Text = "infinity"
Exit Sub
Else
op1 = op1 / Val(Text1.Text)
End If
End If
End If
Text1.Text = ""
opr = "+"
cnt = cnt + 1
End If
lst = lst + " +"
End Sub

Private Sub Command8_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
Exit Sub
Else
If Val(Text1.Text) < 0 Then
Text1.Text = "Imaginary!"
Exit Sub
ElseIf Val(Text1.Text) >= 0 Then
Text1.Text = Val(Text1.Text) ^ (1 / 2)
End If
End If
lst = lst + " <- sqrt is -> " + Text1.Text + "<finish> "
End Sub

Private Sub Command9_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
Exit Sub
Else
Text1.Text = Val(Text1.Text) / 100
End If
lst = lst + "<- % is-> " + Text1.Text + "< finish > "
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Text1.Text = ""
d = 0
op1 = 0
op2 = 0
cnt = 0
lst = ""
Command16.Visible = False
Label2.Visible = False
End Sub


