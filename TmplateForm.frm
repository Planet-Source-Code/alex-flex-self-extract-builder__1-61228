VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Self Extracter"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   Icon            =   "TmplateForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Extract"
      Height          =   735
      Left            =   120
      Picture         =   "TmplateForm.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Files will be extracted in C:\"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' Procedure  : Self extracter
' Created by : Flex
' Date-Time  : 18-6-2005 - 15:09:43
' Description: Self extracter
' License    : See LICENSE.txt
'--------------------------------------------------------------------------------
'This is the template, it's compiled to donotremove.txt

Dim PropBag As New PropertyBag 'This is the other propbag

Private Sub Command1_Click()
Dim FileCount As Long 'We put total files in here.
Dim FileName As String 'Put filename into this
Dim ByteArr() As Byte 'Dim byetarray
FileCount = Val(PropBag.ReadProperty("FileCount"))

For a = 1 To FileCount 'Start the for
FileName = PropBag.ReadProperty("File" & a & "Name") 'Get the filename
Open "C:\" & FileName For Binary As #1 'Open the destination file in binary
ByteArr() = PropBag.ReadProperty("File" & a)  'Get the file

Put #1, , ByteArr() 'Put the whole file in this, it's all!
Close #1 'Close the file

Next a 'Next file

MsgBox "Files extracted,please read license.txt"

'THIS IS ALL, NOW EVERYTHING IS DONE!!!!
'CREATED BY FLEX SOFTWARE PLEASE READ LICENSE.TXT
'
'
'THANKS FOR TRYING THIS

End Sub

Private Sub Form_Load()
Dim BeginPos As Long 'Dim the beginpos to get
Dim tmpByte As Variant 'This is a temp byte arr
Dim ByteArr() As Byte 'Dim byte array

Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1 'Open the .EXE (itself) in binary mode

Get #1, LOF(1) - 3, BeginPos 'Get the start position where the propbag begins (you can see that in de builder project)

Seek #1, BeginPos 'Start pointer at the beginpos
Get #1, , tmpByte 'Get the temp bytes

Close #1 'Close the EXE

ByteArr() = tmpByte 'Put the temp bytes in the real bytes
PropBag.Contents = ByteArr() 'Put the bytes in propertybag, so we can get the files! :D

'END, now click on EXTRACT to extract files
End Sub
