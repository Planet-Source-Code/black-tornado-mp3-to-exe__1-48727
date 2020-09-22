VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWriter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3 to EXE - By Black Tornado"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   Icon            =   "frmWriter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWriter.frx":08CA
   ScaleHeight     =   5670
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowName 
      Caption         =   "Show Name"
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdEditName 
      Caption         =   "Edit Name"
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      Height          =   615
      Left            =   840
      ScaleHeight     =   555
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   360
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select an MP3 File"
      Filter          =   "MP3 Files (*.mp3)|*.mp3"
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdChangeAbout 
      Caption         =   "Change 'About'"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdGUI 
      Caption         =   "Change GUI"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Delete MP3"
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add MP3"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox lstMP3 
      BackColor       =   &H001D1F23&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4050
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7320
      Picture         =   "frmWriter.frx":E541
      ToolTipText     =   "MP3 to EXE, powered by Black Tornado"
      Top             =   4920
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of MP3 Files"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1155
   End
End
Attribute VB_Name = "frmWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Abt As String            ' About TEXT
Dim TheFile$
Dim I As Long                ' Used for FOR ... NEXT
Dim MyBag As New PropertyBag ' Make a new property bag to put in the EXE
Dim MP3_Data$                ' It is very very important that you put the '$'
                             ' after the name of the variable, because if you didn't
                             ' then the data will not be written correctly to the bag
                             ' '$' makes your string accept all letters and symbols
Dim MP3_Info(14) As String    ' Set limit of MP3 files to 15 becaue they are enough!!!
Dim TempVal As String        ' Although it is temprorery, but it is very useful


Private Sub cmdAbout_Click()
MsgBox "MP3 to EXE, by Black Tornado" + vbCrLf & "(C) Copyright 2000-2004 Black Tornado Software" + vbCrLf & _
       "E-mail  : btsoft@burntmail.com" + vbCrLf & _
       "Website : www.BlackTornado.cjb.net" + vbCrLf & _
       "Feel free to send me your comments..."
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Canceled:
If lstMP3.ListCount = 15 Then MsgBox "15 files are too enough, what do you want to do?", vbQuestion: Exit Sub
ComDlg.Filter = "MP3 Files (*.mp3)|*.mp3"
' The value for the property for Comdlg.CancelError = True
' This will help us, if there is an error then it was because user
' canceled, so we will not add the filename property of comdlg, because
' if we chose a file, and we didn't chose another, the FileName property
' still at its last value, wich is the old.
ComDlg.ShowOpen ' Show the open dialog
' Now check if the file exists in the list
For I = 0 To lstMP3.ListCount - 1
If lstMP3.List(I) = ComDlg.FileName Then MsgBox "File already exists in the list", vbExclamation: Exit Sub
Next
lstMP3.AddItem ComDlg.FileName    ' Add item to list
MP3_Info(lstMP3.ListCount - 1) = "Song " & lstMP3.ListCount
Exit Sub ' Mission Successfull
Canceled:
' Don't do any thing!!!
End Sub

Private Sub cmdChangeAbout_Click()
Abt = InputBox("Enter informations displayed in about dialog", , MyBag.ReadProperty("About"))
End Sub

Private Sub cmdCompile_Click()
If lstMP3.ListCount = 0 Then MsgBox "You must add files first!", vbExclamation: Exit Sub
If lstMP3.ListCount = 15 Then MsgBox "15 files are too enough, what do you want to do?", vbQuestion: Exit Sub
On Error GoTo Canceled
ComDlg.FileName = "Test.exe"
ComDlg.Filter = "Executable Files (*.exe)|*.exe"
ComDlg.ShowSave
Screen.MousePointer = vbHourglass
' Begin of code Writer.exe
' The code below, will attempt to mix Reader.exe (PropertyBag reader) with A property bag contents.
' In this example you will see the words EXTRA DATA, these 2 word means Property Bag Contents
' which is really an EXTRA DATA.
' On Error Resume Next ' Resume code even if error happened, I have typed Resume because we will delete
' the file (Test.exe) if exists. But if it is not exists an error will happen
' So we will say 'On Error GoTo HaveError' when we reach the level of code after
' Kill statement, because now every error will happen after the Kill statement
' Its reason is error in opening file (Reader.exe) or in writing test file (Test.exe)
Dim EXE_File As String ' The compiled file path
Dim Writing_Position As Long ' This variable stores the position of writing EXTRA CODE!
Dim Temp As Variant ' This value will store the contents of MyBag, It is a variant because there is
' no specific type of Property Bag contents, it may by picture, string, integer...
' Now we will write some property items to the property bag, because we will retrieve the contents in the other
' program (Reader). And if the property doesn't contains data we will be not able to know if the program (Reader)
' has really read the data.
For I = 0 To lstMP3.ListCount - 1
DoEvents
Open lstMP3.List(I) For Binary As #1
TheFile$ = String(LOF(1), Chr(0))
Get #1, , TheFile$
MyBag.WriteProperty "MP3" & I, TheFile$
MyBag.WriteProperty "MP3_Info" & I, MP3_Info(I)
Close #1
TheFile$ = ""
Next

Continue:
MyBag.WriteProperty "Files", lstMP3.ListCount - 1 ' Number of MP3 Files
MyBag.WriteProperty "Background", Me.Picture
MyBag.WriteProperty "About", Abt
' If you want to add your own property just type:
' MyBag.WriteProperty "Your Property", "Property Value"
' Now, we will make a copy of Reader.exe to a new file named Test.exe
' this copy will read the EXTRA DATA. But there is no EXTRA DATA in it so
' we will open the file as binary and then add our bag contents
EXE_File = ComDlg.FileName
' Kill EXE_File ' Kill the compiled file if it is already exists
FileCopy App.Path & "\Reader.exe", EXE_File ' Make a copy of Reader.exe (Template) to EXE_File

' Now, we are going to write the REAL CODE!!!
On Error GoTo HaveError ' If there is an error, then its a writing error
Open EXE_File For Binary As #1 ' Open template file for applying patch in binary mode
Writing_Position = LOF(1) ' Writing_Position=Length Of File (1) which is the file we opened with Open
' statement (EXE_File), So Writing_Position = Size Of Target File (EXE_File) = End of original file
Temp = MyBag.Contents ' Copy the contents of property bag into Temp variable
Seek #1, LOF(1) ' Set position of file writing to the end of original file (File without EXTRA DATA)
Put #1, , Temp ' Put the contents of property bag to Test.exe file
Put #1, , Writing_Position ' At last, we must put the original file size (length)

Close #1 ' Close the file
Screen.MousePointer = vbDefault
' If operation is done without any errors, that means that the file has been written successfully
MsgBox "File(s) packed successfully, press 'OK' to test the file", vbInformation, "Congratulations!"
Shell EXE_File, vbNormalFocus ' Run the patched file
End ' End the program
HaveError: ' Sorry, we have an error

MsgBox "Error during file compilation", vbCritical, "Error compiling file"
' End of code Writer.exe
 ' 0 ~ Count-1
Canceled:
Exit Sub
End ' End the program
End Sub

Private Sub cmdEditName_Click()
If lstMP3.ListCount = 0 Then Exit Sub
If lstMP3.ListIndex >= 0 Then MP3_Info(lstMP3.ListIndex) = InputBox("Enter new name:", "Enter the current MP3 name", MP3_Info(lstMP3.ListIndex))
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdGUI_Click()
On Error GoTo CancelOperation
ComDlg.Filter = "All Picture File Types|*.bmp;*.gif;*.jpeg;*.jpg;*.wmf;*.emf"
ComDlg.ShowOpen ' Show the open dialog
Me.Picture = LoadPicture(ComDlg.FileName)
CancelOperation:
' Don't do anything
End Sub

Private Sub cmdRemove_Click()
If lstMP3.ListIndex < 0 Then MsgBox "Select an item first", vbCritical: Exit Sub
lstMP3.RemoveItem lstMP3.ListIndex
End Sub

Private Sub cmdShowName_Click()
If lstMP3.ListCount = 0 Then Exit Sub
If lstMP3.ListIndex >= 0 Then MsgBox MP3_Info(lstMP3.ListIndex)
End Sub

Private Sub Form_Load()
' Now, set the flags of ComDlg to a good flags
ComDlg.Flags = cdlOFNHideReadOnly + cdlOFNExplorer
' Now, write default properties
Abt = "MP3-TO-EXE EXAMPLE, CREATED BY BLACK TORNADO"
MyBag.WriteProperty "Background", Me.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Thank you for running MP3-to-EXE, please vote for me if you like this program"
End
End Sub
