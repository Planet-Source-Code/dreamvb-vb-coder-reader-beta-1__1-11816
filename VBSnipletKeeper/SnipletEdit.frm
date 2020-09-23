VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VB Sniplet Editor"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "I am Done"
      Height          =   360
      Left            =   1395
      TabIndex        =   10
      Top             =   5070
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   45
      TabIndex        =   3
      Top             =   3075
      Width           =   6240
      Begin VB.TextBox Text5 
         Height          =   300
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "All"
         Top             =   1425
         Width           =   540
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "10/1/2000"
         Top             =   1080
         Width           =   2970
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   8
         Text            =   "Ben Jones"
         Top             =   690
         Width           =   2970
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "Disable X button on a form"
         Top             =   270
         Width           =   2970
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Vb Ver"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1455
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1095
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Code By"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   750
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code Name"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   840
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Tip"
      Height          =   360
      Left            =   60
      TabIndex        =   1
      Top             =   5070
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "SnipletEdit.frx":0000
      Top             =   630
      Width           =   6255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your code must not exceed over 1.00 KB i size"
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   315
      Width           =   3315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your sniplet code below"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   75
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Code_String As String
Sub SaveToFiles()
    Open App.Path & Module1.Sniplet_Index For Append As #1
    Open App.Path & Module1.Sniplet_Code For Append As #2
    
    Print #1, Text2.Text & "|" & Text3.Text & "|" & Text4.Text & "|" & Text5.Text & "|" & Len(Code_String) & "|"
    Print #2, Code_String
    Close #1
    Close #2
    
End Sub

Private Sub Command1_Click()
Text2 = Trim(Text2)
Text3 = Trim(Text3)
Text4 = Trim(Text4)
Text5 = Trim(Text5)
Code_String = Module1.Encode(Text1.Text)

SaveToFiles
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""

End Sub

Private Sub Command2_Click()
    Unload Form1
    VbReader.Show
    

End Sub

Private Sub Form_Load()
Dim Tindex, TCode As String
Dim IsHere_Index, IsHere_Code As Integer

Module1.CenterForm Form1
    
    Tindex = App.Path & Module1.Sniplet_Index
    TCode = App.Path & Module1.Sniplet_Code
            IsHere_Index = Module1.FileExists(Tindex)
            IsHere_Code = Module1.FileExists(TCode)
            
            If IsHere_Index = 0 Then
                MsgBox "can't find Main index file"
               ElseIf IsHere_Index = 1 And IsHere_Code = 0 Then
                    MsgBox "Can't Find Code File", vbCritical
                    
               Exit Sub
               Else

               End If
               
               
                
            


End Sub

