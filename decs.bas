Attribute VB_Name = "decs"
Option Explicit


' Declaire 2 Varables to Hold the USers name and password and email for the service
Public UserName As String
Public PassWord As String
Public EmailAddress As String
'*
' Varables used for holding serfer time and server date
    Public sysdate As String
    Public systime As String
' API function to execute Objects
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'*

Public Sub LoadAddressBook()
    ' Clear the list box
    Call frmSend.lstAddress.Clear
    ' Set a string for the location of the file
    Dim address As String
   ' Add the location of the file to the string
    address = "C:\AddressBook.txt"
    ' A varable for the loop to read the data from the file into
    Dim n As Variant
    
    Dim infile As Integer
    infile = FreeFile
    ' Open the file to be inputed into the list box
    Open address For Input As #infile
    ' A loop until the end of the file to read the data in
    Do While Not EOF(infile)
        Line Input #infile, n
        'add the data read in to lstAddress
        frmSend.lstAddress.AddItem n
    Loop
    ' close the file
    Close infile
End Sub
Public Sub LoadAddressBook2()
    ' Clear the list box
    Call frmAddress.lstAddress.Clear
    ' Set a string for the location of the file
    Dim address As String
   ' Add the location of the file to the string
    address = "C:\AddressBook.txt"
    ' A varable for the loop to read the data from the file into
    Dim n As Variant
    
    Dim infile As Integer
    infile = FreeFile
    ' Open the file to be inputed into the list box
    Open address For Input As #infile
    ' A loop until the end of the file to read the data in
    Do While Not EOF(infile)
        Line Input #infile, n
        'add the data read in to lstAddress
        frmAddress.lstAddress.AddItem n
    Loop
    ' close the file
    Close infile
End Sub

