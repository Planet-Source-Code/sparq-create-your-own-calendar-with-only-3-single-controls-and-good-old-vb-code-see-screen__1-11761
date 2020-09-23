VERSION 5.00
Begin VB.Form frmChange 
   Caption         =   "Change Calendar"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form2"
   ScaleHeight     =   1365
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   960
      Width           =   1395
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   300
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   435
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   420
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   60
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Month:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cell Height:"
      Height          =   195
      Left            =   2100
      TabIndex        =   1
      Top             =   450
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cell Width:"
      Height          =   195
      Left            =   2100
      TabIndex        =   0
      Top             =   90
      Width           =   765
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.DaySlot(0).Width = Val(Text1)
    Form1.DaySlot(0).Height = Val(Text2)
    Form1.M = Combo1.ListIndex + 1
    Form1.Y = Val(Combo2.Text)
    Form1.SetupDays
End Sub

Private Sub Form_Load()
  Dim TempDate As Date
    Combo1.Clear
    Combo2.Clear
    
    
    For X = 0 To 11
        TempDate = X + 1 & "/01/00"
        Combo1.AddItem Format(TempDate, "mmmm")
    Next X
    Combo1.ListIndex = Month(Date) - 1
    
    For X = 1900 To 2100
        Combo2.AddItem Format(X, "0000")
    Next X
    Combo2.Text = Format(Date, "yyyy")
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
End Sub

