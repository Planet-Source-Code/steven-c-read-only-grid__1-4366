VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Move x"
      Height          =   375
      Left            =   1200
      TabIndex        =   22
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
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
      Left            =   1800
      TabIndex        =   20
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
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
      Left            =   1320
      TabIndex        =   19
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
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
      Left            =   840
      TabIndex        =   18
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
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
      Left            =   360
      TabIndex        =   17
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   14
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   12
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Min             =   1
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   1815
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Min             =   1
      Max             =   100
      Orientation     =   1179648
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Grid1 As GridClass

Private Sub Command1_Click()
Grid1.MoveFirst
End Sub

Private Sub Command2_Click()
Grid1.MovePrevious
End Sub

Private Sub Command3_Click()
Grid1.MoveNext
End Sub

Private Sub Command4_Click()
Grid1.MoveLast
End Sub

Private Sub Command5_Click()
If (Len(Text4.Text) <> 0) And (IsNumeric(Text4.Text)) Then
    Grid1.move Text4.Text
End If
End Sub

Private Sub FlatScrollBar1_Change()
Grid1.ScrollChange
End Sub


Private Sub Form_Load()
Dim iRow As Integer     'current rows
Dim iCol As Integer     'current col
    
    Set Grid1 = New GridClass
    With Grid1
        .Cols = 3
        .Rows = 5
        For iCol = 0 To 2
            .Col = iCol
            Select Case iCol
                Case 0
                    .ColumnDataField = "Au_ID"
                Case 1
                    .ColumnDataField = "Author"
                Case 2
                    .ColumnDataField = "Year Born"
            End Select
            For iRow = 0 To 4
                .Row = iRow
                Select Case iCol
                    Case 0
                        Set .ColumnDisplay = Text1(iRow)
                    Case 1
                        Set .ColumnDisplay = Text2(iRow)
                    Case 2
                        Set .ColumnDisplay = Text3(iRow)
                End Select
            Next iRow
        Next iCol
        Set .ProgressDisplay = ProgressBar1
        Set .ScrollBar = FlatScrollBar1
        .DatabaseName = "biblio.mdb"
        .Recordset = "SELECT * FROM Authors WHERE (((Author) Like 'A*' ))"
    End With
End Sub
