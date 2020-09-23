VERSION 5.00
Begin VB.Form VbReader 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB sniplet Reader beta 1"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "Reader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   -15
      TabIndex        =   32
      Top             =   6150
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   6210
      TabIndex        =   31
      Top             =   5160
      Width           =   720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   345
      Left            =   5205
      TabIndex        =   30
      Top             =   5160
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Sniplet"
      Height          =   345
      Left            =   3660
      TabIndex        =   29
      Top             =   5160
      Width           =   1485
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "Reader.frx":063A
      Left            =   -105
      List            =   "Reader.frx":063C
      TabIndex        =   28
      Top             =   6510
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CommandButton cmdcopy 
      Caption         =   "Copy Sniplet"
      Height          =   345
      Left            =   2100
      TabIndex        =   27
      Top             =   5160
      Width           =   1485
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "<< Return to List"
      Height          =   345
      Left            =   255
      TabIndex        =   26
      Top             =   5160
      Width           =   1770
   End
   Begin VB.TextBox txtdisplaycode 
      Height          =   3375
      Left            =   210
      MaxLength       =   1024
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   25
      Text            =   "Reader.frx":063E
      Top             =   6705
      Visible         =   0   'False
      Width           =   7320
   End
   Begin VB.ListBox lstList 
      Height          =   3375
      Left            =   30
      TabIndex        =   24
      Top             =   1680
      Width           =   7320
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   7035
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   22
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   255
      Index           =   6
      Left            =   6705
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   21
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C000C0&
      Height          =   255
      Index           =   5
      Left            =   6390
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   20
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   6075
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C000&
      Height          =   255
      Index           =   3
      Left            =   5745
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C0C0&
      Height          =   255
      Index           =   2
      Left            =   5430
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   5115
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   1230
      Width           =   255
   End
   Begin VB.ComboBox comboSize 
      Height          =   315
      Left            =   3105
      TabIndex        =   14
      Top             =   1200
      Width           =   660
   End
   Begin VB.ComboBox combofont 
      Height          =   315
      Left            =   480
      TabIndex        =   11
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   30
      TabIndex        =   0
      Top             =   105
      Width           =   7335
      Begin VB.TextBox txtver 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6225
         TabIndex        =   10
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtsize 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4350
         TabIndex        =   8
         Top             =   585
         Width           =   1215
      End
      Begin VB.TextBox txtdate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4350
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtname 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   2640
      End
      Begin VB.TextBox txtBy 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   585
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "VB Ver"
         Height          =   195
         Left            =   5685
         TabIndex        =   9
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Left            =   3915
         TabIndex        =   7
         Top             =   585
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
         Height          =   195
         Left            =   3915
         TabIndex        =   5
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code By:"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code Name:"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Font Colour"
      Height          =   195
      Left            =   3855
      TabIndex        =   23
      Top             =   1275
      Width           =   810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   -195
      X2              =   585
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   -240
      X2              =   540
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Font Size"
      Height          =   210
      Left            =   2310
      TabIndex        =   13
      Top             =   1275
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Font"
      Height          =   195
      Left            =   105
      TabIndex        =   12
      Top             =   1275
      Width           =   315
   End
End
Attribute VB_Name = "VbReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VB Code Reader Coded by Ben Jones

' This is a small program I made for keeping Visual basic Sniplets
' This program also has a inbuilt Editor so you can add your own code
' This code is free to use in any of the programs you make. as long as it's not sold
' If you make any updates to this code I like to know about them
' Email my at Dreamvb@yahoo.com

' Copyright © 1997-2000 Ben Jones

Private Copy_Code_Name As String

Public Function BreakString(TString As String) As String
Dim Tag As Collection
Dim I_Count As Integer
Dim TSplit As String
On Error Resume Next
Set Tag = New Collection

' This is just a small Split function I made splits a string seperated by a |

For I_Count = 1 To Len(TString)
    ch = Mid(TString, I_Count, 1)
    TSplit = TSplit & ch
     If InStr(TSplit, "|") Then
        g = Trim(Left(TSplit, Len(TSplit) - 1)): TSplit = ""
        Tag.Add (g) ' Adds the strings to the Collection
        
    End If
    Next
   
   txtname.Text = Tag(1)
   txtBy.Text = Tag(2)
   txtdate.Text = Tag(3)
   txtver.Text = Tag(4)
   txtsize.Text = Format(Tag(5), "#,#") & " Bytes" ' Just formats the size of the code
   
End Function
Private Sub cmdback_Click()
    txtdisplaycode.Visible = False
    lstList.Visible = True
    cmdcopy.Enabled = False
    cmdback.Enabled = False
    
End Sub

Private Sub cmdcopy_Click()
 On Error Resume Next
    Clipboard.SetText txtdisplaycode.Text
     If Err Then
        MsgBox Err.Description, vbInformation
        Else
            MsgBox Copy_Code_Name & " has been copyed to the clipboard", vbInformation
        End If
        Copy_Code_Name = "" ' Clear out the var
        
End Sub

Private Sub combofont_Click()
    txtdisplaycode.FontName = combofont.Text
    
End Sub

Private Sub comboSize_Click()
    txtdisplaycode.FontSize = Val(comboSize.Text) ' Changes the text font size
    
End Sub
Sub LoadCodeList()
' This will load up the code file

Dim StrBuffer As String
 Open App.Path & Module1.Sniplet_Code For Input As #2
    Do While Not EOF(2)
     Input #2, StrBuffer ' Input the file line by line
     List2.AddItem StrBuffer ' Add each line to the list box
     Loop ' Loop until we reach the end of the file
     Close #2 ' Close file when we are done
     
End Sub
Sub LoadList()
' This will be used to load up the code exaple names and other info

Dim StrList As String
Open App.Path & Module1.Sniplet_Index For Input As #1
    
    Do While Not EOF(1)
    Input #1, StrList ' Input the file line by line
    
    For I = 1 To Len(StrList) ' Get the lenght of the string
        ch = Mid(StrList, I, 1) ' Get each letter
            If ch = "|" Then    ' Get were we need to split
                B = I           ' this is were we take the string we need
                
                lstList.AddItem "[" & lstList.ListCount + 1 & "] " & Left(StrList, B - 1): Exit For ' this just gets the exaple code name
                End If
                Next
                List1.AddItem StrList ' Add the rest of the file to the list box for latter use
    Loop ' Loop until we get to the end of the file
    Close #1
    
End Sub

Private Sub Command1_Click()
    VbReader.Hide ' Hide the forms window
    Form1.Show ' This will be used to display our sniplet editor
    
End Sub

Private Sub Command2_Click()
    about.Show ' Show out about box
    VbReader.Hide ' Hide the window here
    
End Sub

Private Sub Command3_Click()
Dim Answer
    Answer = _
     MsgBox("Do you want to exit the program now", _
     vbYesNo) ' Used to disply the message when exiting program
    
    If Answer = vbNo Then
       Exit Sub ' Don't go eny more doen the code
       Else
        MsgBox "Please Vote if you like this program..", vbInformation ' Disply last message box
       End ' End the program
    End If
       
End Sub

Private Sub Form_Load()
Dim F_Count As Integer

Sniplets_Path = Module1.AddBackSlash(App.Path & "\Sniplets")
    If Module1.FolderExists(Sniplets_Path) = 0 Then
        MsgBox "Can't Find the sniplets folder now createing one now" & vbCrLf & _
        "You can use the sniplet editor to create your tips", vbInformation
        MkDir App.Path & "\Sniplets"
        
        Exit Sub
    Else
    
    End If
    
    For F_Count = 1 To Screen.FontCount - 1 ' Get the number of fonts installed on system
        combofont.AddItem (Screen.Fonts(F_Count)) ' add all the system form to the combo box
        comboSize.AddItem F_Count * 2 ' Add in font sizes
        
    Next
    
    Module1.CenterForm VbReader
        LoadList ' Loads up the index list for the code
        LoadCodeList ' Loads up the main code
        
    combofont.ListIndex = 4
    comboSize.ListIndex = 4
    txtdisplaycode.Top = lstList.Top
    txtdisplaycode.Left = lstList.Left
    txtdisplaycode.Visible = False
    cmdback.Enabled = False
    cmdcopy.Enabled = False
    
End Sub

Private Sub Form_Resize()
    Line1(0).X2 = Form1.Width ' used to give the form a 3d affect line
    Line1(1).X2 = Form1.Width ' used to give the form a 3d affect line
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload VbReader ' Unload the form
    End ' End the program
    
End Sub

Private Sub List1_Click()
    BreakString List1.Text ' This is a call to the break string function
    
End Sub


Private Sub List2_Click()
    txtdisplaycode.Text = Module1.Encode(List2.Text) ' This will decode the code string in the list box
    
End Sub

Private Sub lstList_Click()
List1.ListIndex = lstList.ListIndex
List2.ListIndex = lstList.ListIndex
    
    Copy_Code_Name = lstList.Text ' this line is not realy needed
    BreakString lstList.Text ' This line beraks all the text in to the right order
    
End Sub

Private Sub lstList_DblClick()
    cmdcopy.Enabled = True
    cmdback.Enabled = True
    txtdisplaycode.Visible = True
    
End Sub

Private Sub Picture1_Click(Index As Integer)
    txtdisplaycode.ForeColor = Picture1(Index).BackColor ' Changes the forcolour of the text box
    
End Sub
