VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Calendar"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cool Things"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   60
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2160
      Top             =   2340
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0000
      Height          =   1035
      Left            =   2100
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label Labelss 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Width           =   855
   End
   Begin VB.Label DaySlot 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   615
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public M As Integer
Public Y As Integer
Dim XX As Integer
Dim YY As Integer
  
Private Sub Command1_Click()
    MsgBox "Click ona day to select that day." & vbCrLf & _
           "Try Clicking the DAY OF WEEK Labels to select every day in the month that is that day." & vbCrLf & _
           "Each box's DATE Value is stored in it's TAG property. hold your mouse over it to see its tag."
End Sub

Private Sub DaySlot_Click(Index As Integer)
  Dim X As Integer
    For X = 0 To DaySlot.Count - 1
        If X = Index Then
            DaySlot(X).BackColor = vbBlue
            DaySlot(X).ForeColor = vbWhite
        Else
            DaySlot(X).BackColor = vbWhite
            DaySlot(X).ForeColor = vbBlack
        End If
    Next X
End Sub

Private Sub DaySlot_DblClick(Index As Integer)
    MsgBox DaySlot(Index).Tag
End Sub

Private Sub Form_Load()
    M = Month(Date)
    Y = Year(Date)
    SetupDays
    
    Load frmChange
    frmChange.Top = Top
    frmChange.Left = Left + Width + 120
    frmChange.Text1 = DaySlot(0).Width
    frmChange.Text2 = DaySlot(0).Height
    frmChange.Show
End Sub


Public Sub SetupDays()
  Dim X As Integer
  Dim H As Integer
  Dim W As Integer

  Dim D As String
  Dim TD As Date
    On Error Resume Next
    For X = 0 To 40
        DaySlot(X).Visible = False
    Next X
  
    Dim Spaces As Integer
    DaySlot(0).Left = DaySlot(0).Width
    H = DaySlot(0).Height
    W = DaySlot(0).Width
    YY = DaySlot(0).Top
    XX = DaySlot(0).Left + DaySlot(0).Width
    TD = M & "/01/" & Y
    D = Format(TD, "ddd")
    DaySlot(0).Caption = 1
    Spaces = GetBlanks(D)
    For X = 1 To (DaysIn(M, Y) + (Spaces - 1))
        If X = 6 Or X = 12 Then
            Beep
        End If
        Load DaySlot(X)
        DaySlot(X).Caption = X - (Spaces - 1)
        If Val(DaySlot(X).Caption) < 1 Then
            DaySlot(X).Caption = ""
            DaySlot(X).Visible = False
        Else
            DaySlot(X).Visible = True
            DaySlot(X).Tag = M & "/" & X - (Spaces - 1) & "/" & Format(Y, "00")
            DaySlot(X).ToolTipText = DaySlot(X).Tag
        End If
        DaySlot(X).Left = XX
        DaySlot(X).Top = YY
        DaySlot(X).Height = DaySlot(0).Height
        DaySlot(X).Width = DaySlot(0).Width
       
        XX = XX + W
        If X Mod 7 = 6 Then
            YY = YY + H
            XX = DaySlot(0).Left
        End If
    Next X
    For X = 0 To 6
        Load Labelss(X)
        Labelss(X).Left = DaySlot(X).Left
        Labelss(X).Top = DaySlot(X).Top - Labelss(X).Height
        Labelss(X).Width = DaySlot(X).Width
        Labelss(X).Visible = True
    Next X
    
    Labelss(0) = "Sun"
    Labelss(1) = "Mon"
    Labelss(2) = "Tue"
    Labelss(3) = "Wed"
    Labelss(4) = "Thu"
    Labelss(5) = "Fri"
    Labelss(6) = "Sat"
    Width = Labelss(6).Left + (Labelss(6).Width * 2)
    
    lblTitle.Alignment = vbCenter
    lblTitle.Left = Labelss(0).Left
    lblTitle.Width = Labelss(6).Left + Labelss(6).Width - Labelss(0).Left
    lblTitle.Top = Labelss(0).Top - (lblTitle.Height + 60)
    
    TD = M & "/01/" & Y
    lblTitle = Format(TD, "mmmm") & ", " & Format(TD, "yyyy")
End Sub


Function DaysIn(Mnth As Integer, Yeer As Integer) As Integer
  Dim ThisDate As Date

    ThisDate = Mnth & "/01/" & Yeer
    Do Until Month(DateAdd("D", 1, ThisDate)) <> Mnth
        ThisDate = DateAdd("D", 1, ThisDate)
    Loop
    DaysIn = Day(ThisDate)
End Function

Function GetBlanks(D As String) As Integer
    Select Case D
        Case "Sun"
            GetBlanks = 0
        Case "Mon"
            GetBlanks = 1
        Case "Tue"
            GetBlanks = 2
        Case "Wed"
            GetBlanks = 3
        Case "Thu"
            GetBlanks = 4
        Case "Fri"
            GetBlanks = 5
        Case "Sat"
            GetBlanks = 6
    End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Labelss_Click(Index As Integer)
  Dim X As Integer
    
    For X = 0 To 34
        DaySlot(X).BackColor = vbWhite
        DaySlot(X).ForeColor = vbBlack
    Next X
    
    X = 0
    On Error GoTo Err
    Do
        DaySlot(Index + (X * 7)).BackColor = vbBlue
        DaySlot(Index + (X * 7)).ForeColor = vbWhite
        X = X + 1
    Loop
Err:
End Sub

Private Sub Timer1_Timer()
    If frmChange.Left = Left + Width + 120 Then Exit Sub
    On Error Resume Next
    DoEvents
    Load frmChange
    frmChange.Top = Top
    frmChange.Left = Left + Width + 120
    frmChange.Text1 = DaySlot(0).Width
    frmChange.Text2 = DaySlot(0).Height
    frmChange.Show
    Me.SetFocus
End Sub
