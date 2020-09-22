VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Self Exctract BUILDER"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   1560
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete selected"
      Height          =   855
      Left            =   4680
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add file"
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Build"
      Height          =   615
      Left            =   120
      Picture         =   "Form1.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox FileList 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "(c) Flex Software 2005 - All rights reserved"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Add some files in the list, and click on Build."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'--------------------------------------------------------------------------------
' Procedure  : Self extracter BUILDER
' Created by : Flex
' Date-Time  : 18-6-2005 - 15:09:43
' Description: Self extracter
' License    : See LICENSE.txt
'--------------------------------------------------------------------------------


Private Sub Command1_Click()
Dim PropBag As New PropertyBag 'Make the propertybag, a usefull class
Dim ByteArr() As Byte 'Dim the byte array we use to put the files in and then in the propertybag
FileCopy App.Path & "\donotremove.txt", "C:\selfextract.exe" 'Copy the Install exe, it's a template exe where we put the files in

PropBag.WriteProperty "FileCount", FileList.ListCount 'Notify the self extracter how many files we have.

'Now we are going to put the files in the propertybag
Dim a As Integer
For a = 0 To FileList.ListCount - 1

Open FileList.List(a) For Binary As #1 'Open the file, in binary
ReDim ByteArr(0 To LOF(1) - 1) 'Redim the bytearray
Get #1, , ByteArr() 'Get the file in our memory
Close #1 'Close the file

PropBag.WriteProperty "File" & a + 1, ByteArr() 'Put the file in the propertybag, to get it back in the self extract exe.
PropBag.WriteProperty "File" & a + 1 & "Name", Mid(FileList.List(a), InStrRev(FileList.List(a), "\") + 1) 'Put the original filename, so the self extracter knows how to name the filename.
Next a

'We are done with puttin the files in the propertybag! It's so easy! Just 5-10 lines of code!





'Now we need to put the propertybag in the template exe:
Open "C:\selfextract.exe" For Binary As #1 'Open the copied exe, in binary mode.
Dim BeginPos As Long 'Dim variable to use for beginpos
BeginPos = LOF(1) 'Get the total length of the file
Seek #1, LOF(1) 'Start pointer at end of the file
Put #1, , PropBag.Contents 'Put the whole propertybag with all the files in the template exe, so easy!
Put #1, , BeginPos 'Add the long where the propertybag starts with it's data, so the self extracter knows it.
Close #1 'Close the file.

MsgBox "Done!"
'NOW YOU MADE YOUR VERY OWN SELF EXTRACTER!

'Please see Template.vbp how that works, because this isn't all!
End Sub

Private Sub Command2_Click()
cd.ShowOpen 'Open the file select dialog
If cd.FileName <> vbNullString Then 'Check if the user didn't choose X
FileList.AddItem cd.FileName 'Add the file into the list
End If
End Sub

Private Sub Command3_Click()

If FileList.ListCount > 0 Then 'Check if the list contains files
FileList.RemoveItem FileList.ListIndex 'Delete the selected file
End If
End Sub
