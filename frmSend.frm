VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSend 
   Caption         =   "Materialised SMSX SMS Sender Ver 1.0"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   2655
      Left            =   360
      TabIndex        =   16
      Top             =   6120
      Width           =   5655
      ExtentX         =   9975
      ExtentY         =   4683
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CheckBox chkConfirm 
      Caption         =   "Would You Like Confirmation of the Message being sent?"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtMessage 
      Height          =   1215
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox txtDeliver 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   4800
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   873
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9000
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
      ExtentX         =   2143
      ExtentY         =   873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   9120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9600
      Top             =   1440
   End
   Begin VB.TextBox txtPhoneNumber 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Picture         =   "frmSend.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdAddress 
      Caption         =   "Load Number Book"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      Picture         =   "frmSend.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdManage 
      Caption         =   "Number Manager"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Picture         =   "frmSend.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox lstAddress 
      Height          =   4545
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Your Name"
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Delivery Time"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Number to Send to"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddress_Click()
Call LoadAddressBook

End Sub

Private Sub cmdManage_Click()
' Load the address manager
    Call Load(frmAddress)
    Call frmAddress.Show
End Sub

Private Sub cmdSend_Click()
    If txtName.Text = "" Then
        MsgBox "Please enter your Nam as this will appear on the recipiants phone, Please only enter your first name", 16, "Error"
        txtName.SetFocus
        Exit Sub
    End If
    
    If txtDeliver.Text = "" Then
        MsgBox "The server relies on being sent a delivery time, please correct this problam", 16, "Error"
        txtDeliver.SetFocus
        Exit Sub
    End If
    
    If txtPhoneNumber.Text = "" Then
        MsgBox "How do you expect to send a text message if you dont enter a valid Number?", 16, "Error"
        txtPhoneNumber.SetFocus
        Exit Sub
    End If
    
    If chkConfirm.Value = 1 Then
        EmailAddress = InputBox("Please Enter the email Account you would like confirmation delivered to", "Enter Email Address")
    
            If Len(EmailAddress) = 0 Then
                MsgBox "You Selected confirmation, and did not enter a valid email address, Please resend your message using a correct email address.", 16, "Error"
                Exit Sub
            End If
    End If
    
    If chkConfirm.Value = 0 Then
        EmailAddress = "ThisIsMadeUp@here.com"
    End If
    Dim ToSend As String
    ToSend = txtMessage.Text + "%3F"
    ToSend = Replace(ToSend, " ", "+")
    
    
    WebBrowser2.Navigate ("http://www.soapengine.com/lucin/soapenginex/smsx.asmx/SendMessage?UserName=" & UserName & "&PassKey=" & PassWord & "&PhoneNumber=" & txtPhoneNumber.Text & "&Message=" & ToSend & "&SenderName=" & txtName.Text & "&ScheduledDate=" & sysdate & "&ScheduledTime=" & txtDeliver.Text & "&Reference=" & "" & "&Confirmation=" & EmailAddress & "&UseFlash=0")
    
End Sub

Private Sub Form_Load()
'*********** Settings For Address Book ***********
 ' If the file exists, do nothing
     If Dir("c:\addressbook.txt") <> "" Then
        DoEvents
    Else
        ' if it doesnt exist, create the address book file
        Dim FilePath As String
        FilePath = "c:\addressbook.txt"
        Open FilePath For Output As #3
        ' close the file
        Close #3
    End If
'***************************************************
Timer1.Enabled = True
End Sub

Private Sub lstAddress_Click()
    Dim blank As Integer
    Dim transfer As String
    Dim final As String
    ' Assign the value clicked in the list box to a varable so we can minipulate it
    transfer = (lstAddress.List(lstAddress.ListIndex))
    ' Use InStr to find the first space in the string, and assign its location to blank
    blank = InStr(1, transfer, " ")
    ' Assign the string from transfer to the varable final, triming it by the value produced in blank and -1 to delete the space
    final = Left$(transfer, blank - 1)
    ' Transfer the formated phone number to the text box
    txtPhoneNumber.Text = final
End Sub


Private Sub Timer1_Timer()
'*********** Get the Server Date & Time  ***********
    Text1.Text = Inet1.OpenURL("http://www.soapengine.com/lucin/soapenginex/smsx.asmx/GetServerTime?UserName=" & UserName & "&PassKey=" & PassWord)
    Dim Combined As String
    
    ' Assign the recived data from the server to a text box
    Combined = Text1.Text
    ' Trim off the XML formating
    Combined = Right(Combined, 28)
    ' Obain the System date from the formated String
    sysdate = Left(Combined, 10)
    ' Now we need to get the system time from the text box
    ' so we isolate the last 17 characters and assign them to varable combined
    Combined = Right$(Text1.Text, 17)
    ' We know the time is 5 characters long, so we simply take them, and assign them to the
    ' varable systime
    systime = Left(Combined, 5)
    ' Enable the timer to update the varables every second
    
'***************************************************
StatusBar1.SimpleText = "The Current Server Time is Now " & systime & " On " & sysdate

End Sub

Private Sub txtDeliver_GotFocus()
MsgBox "Make sure you enter a time in the future to the Current Server Time. The current Server time is " & systime, vbInformation + vbOKOnly, "Information"

End Sub
