VERSION 5.00
Begin VB.Form frmIPTools 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IPTools - by Jamie Frost 2002"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   Icon            =   "frmIPTools.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   330
      Left            =   105
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1050
      Width           =   750
   End
   Begin VB.Frame frmNetwork 
      Caption         =   "Network"
      Height          =   960
      Left            =   105
      TabIndex        =   41
      Top             =   3780
      Width           =   3900
      Begin VB.Frame frmID 
         Caption         =   "ID"
         Height          =   645
         Left            =   105
         TabIndex        =   44
         Top             =   210
         Width           =   2220
         Begin VB.TextBox txtNwIdDec 
            Height          =   330
            Index           =   4
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   210
            Width           =   435
         End
         Begin VB.TextBox txtNwIdDec 
            Height          =   330
            Index           =   3
            Left            =   1155
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   210
            Width           =   435
         End
         Begin VB.TextBox txtNwIdDec 
            Height          =   330
            Index           =   2
            Left            =   630
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   210
            Width           =   435
         End
         Begin VB.TextBox txtNwIdDec 
            Height          =   330
            Index           =   1
            Left            =   105
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   45
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   210
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "."
            Height          =   225
            Index           =   15
            Left            =   1575
            TabIndex        =   51
            Top             =   315
            Width           =   120
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "."
            Height          =   225
            Index           =   16
            Left            =   1050
            TabIndex        =   50
            Top             =   315
            Width           =   120
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "."
            Height          =   225
            Index           =   17
            Left            =   525
            TabIndex        =   49
            Top             =   315
            Width           =   120
         End
      End
      Begin VB.Frame frmClass 
         Caption         =   "Class"
         Height          =   645
         Left            =   2520
         TabIndex        =   42
         Top             =   210
         Width           =   1275
         Begin VB.TextBox txtClass 
            Height          =   330
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "Class: N/A"
            Top             =   210
            Width           =   1065
         End
      End
   End
   Begin VB.CommandButton cmdDefaultSubnetMask 
      Caption         =   "Default"
      Height          =   330
      Left            =   3255
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1050
      Width           =   750
   End
   Begin VB.Frame frmNetworkIDBin 
      Caption         =   "Network ID in Binary (result of anding)"
      Height          =   645
      Left            =   105
      TabIndex        =   32
      Top             =   3045
      Width           =   3900
      Begin VB.TextBox txtNwIdBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   1
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox txtNwIdBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   4
         Left            =   2940
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox txtNwIdBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   3
         Left            =   1995
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox txtNwIdBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   2
         Left            =   1050
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   14
         Left            =   2835
         TabIndex        =   39
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   13
         Left            =   1890
         TabIndex        =   38
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   12
         Left            =   945
         TabIndex        =   37
         Top             =   315
         Width           =   120
      End
   End
   Begin VB.Frame frmSubnetMaskBin 
      Caption         =   "Subnet Mask in Binary"
      Height          =   645
      Left            =   105
      TabIndex        =   24
      Top             =   2310
      Width           =   3900
      Begin VB.TextBox txtSubnetBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   2
         Left            =   1050
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "11111111"
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox txtSubnetBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   3
         Left            =   1995
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "11111111"
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox txtSubnetBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   4
         Left            =   2940
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox txtSubnetBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   1
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "11111111"
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   11
         Left            =   945
         TabIndex        =   31
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   10
         Left            =   1890
         TabIndex        =   30
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   9
         Left            =   2835
         TabIndex        =   29
         Top             =   315
         Width           =   120
      End
   End
   Begin VB.Frame frmSubnetMaskDec 
      Caption         =   "Subnet Mask"
      Height          =   645
      Left            =   945
      TabIndex        =   20
      Top             =   840
      Width           =   2220
      Begin VB.TextBox txtSubnetDec 
         Height          =   330
         Index           =   1
         Left            =   105
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "255"
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox txtSubnetDec 
         Height          =   330
         Index           =   2
         Left            =   630
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "255"
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox txtSubnetDec 
         Height          =   330
         Index           =   3
         Left            =   1155
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "255"
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox txtSubnetDec 
         Height          =   330
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         Top             =   210
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   8
         Left            =   525
         TabIndex        =   23
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   7
         Left            =   1050
         TabIndex        =   22
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   6
         Left            =   1575
         TabIndex        =   21
         Top             =   315
         Width           =   120
      End
   End
   Begin VB.Frame frmIPBin 
      Caption         =   "Original IP in Binary"
      Height          =   645
      Left            =   105
      TabIndex        =   12
      Top             =   1575
      Width           =   3900
      Begin VB.TextBox txtIPBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   1
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox txtIPBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   4
         Left            =   2940
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox txtIPBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   3
         Left            =   1995
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox txtIPBin 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   2
         Left            =   1050
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   5
         Left            =   2835
         TabIndex        =   19
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   4
         Left            =   1890
         TabIndex        =   18
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   3
         Left            =   945
         TabIndex        =   17
         Top             =   315
         Width           =   120
      End
   End
   Begin VB.Frame frmIPDec 
      Caption         =   "Original IP"
      Height          =   645
      Left            =   945
      TabIndex        =   8
      Top             =   105
      Width           =   2220
      Begin VB.TextBox txtIPDec 
         Height          =   330
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "0"
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox txtIPDec 
         Height          =   330
         Index           =   3
         Left            =   1155
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "0"
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox txtIPDec 
         Height          =   330
         Index           =   2
         Left            =   630
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox txtIPDec 
         Height          =   330
         Index           =   1
         Left            =   105
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "0"
         Top             =   210
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   2
         Left            =   1575
         TabIndex        =   11
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   1
         Left            =   1050
         TabIndex        =   10
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "."
         Height          =   225
         Index           =   0
         Left            =   525
         TabIndex        =   9
         Top             =   315
         Width           =   120
      End
   End
End
Attribute VB_Name = "frmIPTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Sets the subnet mask to the industry standard for each class of IP
'Same basic thing happens in the IP class determination sub
Private Sub cmdDefaultSubnetMask_Click()

  Dim i As Integer
  Dim intTop As Integer

    'Determine what class the IP is from the value of the first
    'octet
    Select Case Val(txtIPDec(1).Text)
      Case 0 To 126
        intTop = 1
      Case 127
        intTop = 1
      Case 128 To 191
        intTop = 2
      Case Else
        intTop = 3
    End Select

    For i = 1 To 4            'set all to zero
        txtSubnetDec(i).Text = 0
    Next i
    For i = 1 To intTop       'set appropriate to 255
        txtSubnetDec(i) = 255
    Next i
    'Update the boxes for the Subnet Mask Binary
    UpdateSubnetBin

End Sub

'Set the values of the Binary original IP boxes
Private Sub UpdateIpBin()

  Dim i As Integer

    For i = 1 To 4
        txtIPBin(i).Text = DecToBin(Val(txtIPDec(i).Text))
    Next i
    UpdateNwIdBin   'Update the Binary of the network ID

End Sub

'Set the values of the Binary Subnet Mask
Private Sub UpdateSubnetBin()

  Dim i As Integer

    For i = 1 To 4
        txtSubnetBin(i).Text = DecToBin(Val(txtSubnetDec(i).Text))
    Next i
    UpdateNwIdBin   'Update the Binary of the network ID

End Sub

'This is where the actual 'stuff' happens, where the 'anding' occurs
'Anding is the process of binary addition, with a similar output to
'"or":
'1 + 1 = 1
'1 + 0 = 0
'0 + 0 = 0
Private Sub UpdateNwIdBin()

  Dim i As Integer     'for loop variables
  Dim j As Integer     ' ^
  Dim temp As String   'string to hold binary number

    For i = 1 To 4
        temp = "00000000"
        For j = 1 To 8
            If (Mid(txtIPBin(i).Text, j, 1) = Mid(txtSubnetBin(i), j, 1)) And (Mid(txtSubnetBin(i), j, 1) <> "0") Then
                Mid(temp, j, 1) = "1"
              Else
                Mid(temp, j, 1) = "0"
            End If
        Next j
        txtNwIdBin(i).Text = temp
    Next i
    UpdateNwIdDec

End Sub

'Simple binary to decimal conversion for the Network ID decimal form
Private Sub UpdateNwIdDec()

  Dim i As Integer
  Dim j As Integer
  Dim BitVal As Integer
  Dim TotValue As Integer

    For i = 1 To 4
        TotValue = 0
        For j = 1 To 8
            BitVal = 2 ^ (8 - j)
            If Mid(txtNwIdBin(i).Text, j, 1) = "1" Then
                TotValue = TotValue + BitVal
            End If
        Next j
        txtNwIdDec(i).Text = TotValue
    Next i

End Sub

'Updates the IP Class textbox with an appropriate value
Private Sub UpdateIpClass()

  Dim ClassLetter As String

    Select Case Val(txtIPDec(1).Text)
      Case 1 To 126
        ClassLetter = "A"
      Case 127
        ClassLetter = "Local"
      Case 128 To 191
        ClassLetter = "B"
      Case 192 To 223
        ClassLetter = "C"
      Case 224 To 240
        ClassLetter = "D"
      Case 241 To 255
        ClassLetter = "E"
      Case Else
        ClassLetter = "N/A"
    End Select
    txtClass.Text = "Class: " & ClassLetter

End Sub

'I think I covered all my bases with users typing in stupid values,
'whenever the user types almost anything, it gets checked and
'doublechecked for stupidity.
'Also, it highlights the text IN the box when a box is tabbed to.
Private Sub txtIPDec_Change(Index As Integer)

    UpdateIpBin
    'Remove any "."s from the box, as "." is used for tab
    txtIPDec(Index).Text = Replace(txtIPDec(Index), ".", "")
    'Check for letters whenever the box is nonzero length
    If (Not IsNumeric(txtIPDec(Index).Text)) And (Len(txtIPDec(Index).Text) > 0) Then
        MsgBox "Please enter only integer values", vbExclamation, "Entry Error"
        txtIPDec(Index).Text = "0"
        txtIPDec(Index).SelStart = 0
        txtIPDec(Index).SelLength = Len(txtIPDec(Index).Text)
    End If
    'Check for invalid IP numbers (outside 0-255)
    If (Val(txtIPDec(Index).Text) > 255) Or (Val(txtIPDec(Index).Text) < 0) Then
        MsgBox "Please enter only values between 0 and 255 inclusive", vbExclamation, "Entry Error"
        txtIPDec(Index).Text = "0"
        txtIPDec(Index).SelStart = 0
        txtIPDec(Index).SelLength = Len(txtIPDec(Index).Text)
    End If
    UpdateIpClass

End Sub

'Used for generating my own 'taborder'
Private Sub txtIPDec_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = 110 Then 'if they hit "." decimal, tab to the next field
        If Index < 4 Then
            txtIPDec(Index + 1).SetFocus
          Else
            txtSubnetDec(1).SetFocus
        End If
    End If

End Sub

'Set to zero if feild is left blank
Private Sub txtIPDec_LostFocus(Index As Integer)

    If txtIPDec(Index).Text = "" Then
        txtIPDec(Index).Text = "0"
    End If

End Sub

'Same as IP Decimal
Private Sub txtSubnetDec_Change(Index As Integer)

    UpdateSubnetBin
    'Remove any "."s from the box, as "." is used for tab
    txtSubnetDec(Index).Text = Replace(txtSubnetDec(Index), ".", "")
    'Check for letters whenever the box is nonzero length
    If (Not IsNumeric(txtSubnetDec(Index).Text)) And (Len(txtSubnetDec(Index).Text) > 0) Then
        MsgBox "Please enter only integer values", vbExclamation, "Entry Error"
        txtSubnetDec(Index).Text = "0"
        txtSubnetDec(Index).SelStart = 0
        txtSubnetDec(Index).SelLength = Len(txtSubnetDec(Index).Text)
    End If
    'Check for invalid IP numbers (outside 0-255)
    If (Val(txtSubnetDec(Index).Text) > 255) Or (Val(txtSubnetDec(Index).Text) < 0) Then
        MsgBox "Please enter only values between 0 and 255 inclusive", vbExclamation, "Entry Error"
        txtSubnetDec(Index).Text = "0"
        txtSubnetDec(Index).SelStart = 0
        txtSubnetDec(Index).SelLength = Len(txtSubnetDec(Index).Text)
    End If

End Sub

'Used for generating my own 'taborder'
Private Sub txtSubnetDec_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = 110 Then   'If they hit ".", tab to the next field
        If Index < 4 Then
            txtSubnetDec(Index + 1).SetFocus
          Else
            txtIPDec(1).SetFocus
        End If
    End If

End Sub

'Set to zero if feild is left blank
Private Sub txtSubnetDec_LostFocus(Index As Integer)

    If txtSubnetDec(Index).Text = "" Then
        txtSubnetDec(Index).Text = "0"
    End If

End Sub

'Select the text in the box whenever the box gains focus
Private Sub txtIPDec_GotFocus(Index As Integer)

    txtIPDec(Index).SelStart = 0
    txtIPDec(Index).SelLength = Len(txtIPDec(Index).Text)

End Sub

'Select the text in the box whenever the box gains focus
Private Sub txtSubnetDec_GotFocus(Index As Integer)

    txtSubnetDec(Index).SelStart = 0
    txtSubnetDec(Index).SelLength = Len(txtSubnetDec(Index).Text)

End Sub

'Function returns a string, length 8, of the byte passed to it
'converted to binary.
'It goes thru the length of the string, and if that particular bit 'fits'
'in the decimal number, it sets that bit, and subtracts its amount from the decimal
'number...pretty simple.  Theres probably a better way, but I like strings. :)
Private Function DecToBin(ByVal Dec As Integer) As String

  Dim i As Integer
  Dim BitVal As Integer '128,64,32...etc

    DecToBin = "00000000"
    For i = 1 To 8
        BitVal = 2 ^ (8 - i)
        If Dec >= BitVal Then
            Dec = Dec - BitVal
            Mid(DecToBin, i, 1) = "1"
        End If
    Next i

End Function

'duh
Private Sub cmdAbout_Click()

    MsgBox "Well, this is it - nothing too special really..." & vbCrLf & vbCrLf & vbCrLf & "-Jamie", vbInformation, "About"

End Sub

'Set the ToolTipText's for each textbox, not essential, but more professional...
'at least, it is the way I see it
Private Sub Form_Load()

  Dim i As Integer

    For i = 1 To 4
        txtIPDec(i).ToolTipText = "IP octet " & Trim(Str(i))
        txtSubnetDec(i).ToolTipText = "Subnet octet " & Trim(Str(i))
        txtIPBin(i).ToolTipText = "IP octet " & Trim(Str(i)) & " in binary"
        txtSubnetBin(i).ToolTipText = "Subnet Mask octet " & Trim(Str(i)) & " in binary"
        txtNwIdBin(i).ToolTipText = "Network ID octet " & Trim(Str(i)) & " in binary"
        txtNwIdDec(i).ToolTipText = "Network ID " & Trim(Str(i))
        txtClass.ToolTipText = "IP Class"
    Next i

End Sub

'Btw, thanks Ulli, great formatter :)

':) Ulli's VB Code Formatter V2.13.6 (8/22/2002 1:58:43 AM) 1 + 286 = 287 Lines
