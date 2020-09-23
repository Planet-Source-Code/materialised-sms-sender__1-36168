VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Materialised's SMS Sender"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help I'm Stuck"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Text            =   "Enter Your SMSX Password"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Text            =   "Enter Your SMSX Username"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   2160
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLogin.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    End ' Exit the Program
End Sub

Private Sub cmdHelp_Click()
    ' To Do
End Sub

Private Sub cmdLogin_Click()
    ' Make Sure the User entered data
    If txtUsername = "" Then
        MsgBox "You Must Enter a Valid SMSX Username", 16, "Error"
        Exit Sub
    End If
    
    If txtUsername = "Enter Your SMSX Username" Then
        MsgBox "You Must Enter a Valid SMSX USername", 16, "Error"
        Exit Sub
    End If
    If txtPassword.Text = "" Then
        MsgBox "You Must Enter a Valid SMSX Password", 16, "Error"
        Exit Sub
    End If
    If txtPassword.Text = "Enter Your SMSX Password" Then
        MsgBox "You Must Enter a Valid SMSX Password", 16, "Error"
        Exit Sub
    End If
    ' Assign the Inputed data to the 2 varables in decs.bas
    UserName = txtUsername.Text
    PassWord = txtPassword.Text
    Debug.Print UserName
    Debug.Print PassWord
    ' Load and Show The Send Form
    Call Load(frmSend)
    Call frmSend.Show
    ' Unload Login Form
    Unload Me
    
End Sub

Private Sub cmdRegister_Click()
    ' Use Windows API to open the webbrowser on page specified below
    ShellExecute Me.hwnd, "open", "http://www.webservicebuy.com/x/smsreg.asp", "", "", 10
End Sub

Private Sub txtPassword_GotFocus()
    ' Clear the Password Text box When the control recives focus from the user
    txtPassword.Text = ""
    ' Set the .PasswordChar property of the TextBox to the default Windows password Mask
    txtPassword.PasswordChar = "*"
End Sub

Private Sub txtUsername_GotFocus()
    ' Clear the Username Text box When the control recives focus from the user
    txtUsername = ""
End Sub
