VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmReader 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Black Tornado - MP3 to EXE"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   ClipControls    =   0   'False
   Icon            =   "frmReader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmReader.frx":08CA
   ScaleHeight     =   4290
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About MP3-2-EXE"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdAboutA 
      Caption         =   "About Author"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin MCI.MMControl MciCtrl 
      Height          =   450
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   794
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.ListBox lstMP3 
      BackColor       =   &H001D1F23&
      ForeColor       =   &H0000FFFF&
      Height          =   2400
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label lblFileTitle 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select file!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   5175
   End
End
Attribute VB_Name = "frmReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Extracted_Bag As New PropertyBag ' Name of the extracted property bag
Dim Reading_Position As Long ' Start point of file reading
Dim Temp As Variant ' Variable to store the property bag contents from the Extracted_Bag
Dim RealContents() As Byte
Dim File$
Dim I As Long

Private Sub cmdAbout_Click()
MsgBox "MP3 to EXE, by Black Tornado" + vbCrLf & "(C) Copyright 2000-2004 Black Tornado Software" + vbCrLf & _
       "E-mail  : btsoft@burntmail.com" + vbCrLf & _
       "Website : www.BlackTornado.cjb.net" + vbCrLf & _
       "Feel free to send me your comments..."
End Sub

Private Sub cmdAboutA_Click()
MsgBox Extracted_Bag.ReadProperty("About"), vbInformation
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
On Error GoTo ReadError ' If there is an error in reading data, tell the user that!

Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1 ' Open the patched file (the file itself)
Get #1, LOF(1) - 3, Reading_Position ' Gets that start position of data and put it in Reading_Position

Seek #1, Reading_Position ' Set the file read position which is length of file - 3 characters
Get #1, , Temp            ' Temp = PropertyBag.Contents
RealContents = Temp       ' RealData = Temp converted into Bytes

Extracted_Bag.Contents = RealContents ' Put the contents to the bag
Close #1                  ' After we finished reading, we must close the file.
For I = 0 To Extracted_Bag.ReadProperty("Files")
lstMP3.AddItem Extracted_Bag.ReadProperty("MP3_Info" & I)
Next
Me.Picture = Extracted_Bag.ReadProperty("Background")
Exit Sub
ReadError: ' There is an error in reading data from file
           MsgBox "Error during data read, data may be empty or it is invalid", vbCritical
End Sub

Sub ExtractMP3(Index As Integer)
MciCtrl.Notify = False
MciCtrl.Wait = True
MciCtrl.Command = "close"
On Error Resume Next
' Kill the old file if already exists
Kill "C:\MTE.mp3"
File$ = Extracted_Bag.ReadProperty("MP3" & Index)
Open "C:\MTE.mp3" For Binary As #1
Put #1, , File$
Close #1
MciCtrl.FileName = "C:\MTE.mp3"
MciCtrl.Notify = False
MciCtrl.Wait = True
MciCtrl.Command = "open"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
MciCtrl.Command = "close"
Kill "C:\MTE.mp3"
End
End Sub

Private Sub lstMP3_Click()
lblFileTitle.Caption = lstMP3.List(lstMP3.ListIndex)
ExtractMP3 lstMP3.ListIndex
End Sub
