VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   12360
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   5640
      Width           =   5655
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   4920
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rename"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   4440
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   3960
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   3480
      Width           =   5655
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   6840
      Pattern         =   "*.nds"
      TabIndex        =   2
      Top             =   1800
      Width           =   5775
   End
   Begin VB.DirListBox Dir1 
      Height          =   11340
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6615
   End
   Begin VB.DriveListBox Drv1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Name File1.Path & "\" & Text1.Text As File1.Path & "\" & Text2.Text

Name File1.Path & "\" & Text3.Text As File1.Path & "\" & Text4.Text

Text5.Text = Text2.Text
Dir1 = "F:\Nintendo\Roms\Nintendo DS Roms"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

File1.Refresh
Dir1.SetFocus

End Sub

Private Sub dir1_Change()
  File1 = Dir1

End Sub

Private Sub Dir1_Click()
  Text6.Text = Dir1.List(Dir1.ListIndex)
  File1 = Dir1.List(Dir1.ListIndex)
  
End Sub

Private Sub drv1_Change()
  Dir1 = Drv1
  File1 = Dir1
End Sub

Private Sub File1_Click()

Text1.Text = File1

myName = Right(Text1, Len(Text1) - 7)
myNameB = Left(myName, Len(myName) - 4)
myNameC = Replace(myNameB, "(U)", "")
myNameD = Replace(myNameC, "(Legacy)", "")
myNameE = Replace(myNameD, "(Trashman)", "")
myNameF = Replace(myNameE, "(Mode 7)", "")
myNameG = Replace(myNameF, "(Lube)", "")

myNum = Left(Text1, 4)

Text3.Text = myNum & ".jpg"
Text4.Text = Trim(myNameG) & ".jpg"

Text2.Text = Trim(myNameG) & ".nds"

End Sub

Private Sub Form_Load()
Drv1 = "F:"
Dir1 = "F:\Nintendo\Roms\Nintendo DS Roms"
End Sub

