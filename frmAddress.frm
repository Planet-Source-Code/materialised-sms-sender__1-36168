VERSION 5.00
Begin VB.Form frmAddress 
   Caption         =   "Address Book"
   ClientHeight    =   2370
   ClientLeft      =   1455
   ClientTop       =   3720
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   2505
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Picture         =   "frmAddress.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      Picture         =   "frmAddress.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox lstaddress 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdadd_Click()
    Dim AddNum As String
    ' Ask the user for some Input
    AddNum = InputBox("Please enter the number you wish to be added to the Contact list. e.g 447846332134 -Name (There must be a space)", "Input New Number", "1111111111111 -My Girlfriend")
    ' Check Data was Entered
    If Len(AddNum) = 0 Then
        Call Error
        Exit Sub
    End If
    Dim FilePath As String
    ' Set the file path
    FilePath = "c:\addressbook.txt"
    ' Open the file so we can write to it
    Open FilePath For Append As #2
    ' Print the user input into the file
    Print #2, AddNum
    ' Close the file
    Close #2
    
End Sub

Private Sub cmdRefresh_Click()
    ' Look at Functions.bas
    Call LoadAddressBook2
End Sub

