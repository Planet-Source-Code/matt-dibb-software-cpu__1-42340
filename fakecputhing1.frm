VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Fake CPU thing"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "View ROM"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Opcodes?"
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Clear"
      Height          =   255
      Left            =   6000
      TabIndex        =   16
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Execute"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"fakecputhing1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox RamT 
      Height          =   285
      Index           =   5
      Left            =   4680
      TabIndex        =   11
      Text            =   "0"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox RamT 
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   10
      Text            =   "0"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Decode"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox RamT 
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   3
      Text            =   "0"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox RamT 
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Text            =   "0"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox RamT 
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      Text            =   "3"
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox RamT 
      Height          =   285
      Index           =   0
      Left            =   4680
      TabIndex        =   0
      Text            =   "2"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "F"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "E"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "D"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "C"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "B"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "A"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Ram:"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quick and Dirty implementation of a sort of crappy fake CPU
'Not much you can do, buts it just a proof of concept!
'
'The idea is this is a 'software' implementation of a simple
'cpu.  Its based on the usual von Neumann architecture. You
'programme it using the opcodes, which are then converted and
'stored in the rom. When exectuted, the 'cpu' will go
'through the memory and run the programme.
'
'Have fun with it!
'
'Copyright Matt Dibb 2003

'Internal Variables & crap
'==============================================================
Private Stopped As Boolean      'are we running?
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_GETLINE = &HC4
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINEFROMCHAR = &HC9


'Memory hardware
'==============================================================
'byte of memory
Private Type MemByte
    Data As String          'ick i know, need hex though
End Type
'ROM array (15 bytes)
Private Rom(15) As MemByte
'RAM array (4 bytes)
Private Ram(5) As MemByte

'Registers
'==============================================================
Private Type Register
    Data As String
End Type
Private XReg As Register    'X register
Private Acc As Register     'accumulator
Private Res As Register     'result register

'Menomics info
'==============================================================
Private Const hlt = 0   'halt
Private Const lda = 1   'load into acc
Private Const ldx = 2   'load into x
Private Const adx = 3   'add x to acc
Private Const sbx = 4   'sub x from acc
Private Const sta = 5   'store res to address

Private Sub Check1_Click()
'overwrite everything with 0
For r = 0 To 5
    RamT(r).Enabled = True
    RamT(r).Text = "0"
Next r
'return the check to unchecked
Check1.Value = 0
End Sub


Private Sub Command1_Click()
'Now the dirty part! :-)
'We have to look at the code and decide what it will
'be and enter this directly into the rom for execution later.

'We only have 16 bytes of ROM so we are limited, but then
'we only have a couple of registers and not many opcodes ;-)

'Once we have done this, we will then run it on our
'virtual machine (VM) - i.e. this app :-)  Dodgy I know, but
'its only a proof of concept!   I don't know any x86 assembler
'so these codes are just made up ones for the cpu in this VM, so
'dont even think about port this output to a compiler! :-)

'********
'Warning!
'********
'There is no error checking to see if we have filled the ROM,
'so watch out.  Also if it doesn't understand an opcode, it
'will probably just ignore it. Remember, this is just a
'proof of concept!


Dim Lines As Integer
Dim i As Integer
Dim LineData As String
Dim MemoryTemp As String
Dim RomCounter As Integer

Command2.Enabled = False

RomCounter = 0  'we incrememnt this each time to make sure the data is in the right memory space

'Get lines
Lines = SendMessage(Text1.hwnd, EM_GETLINECOUNT, 0, 0)

'Set the cursor back to the start
Text1.SelStart = 1

'save the file because VB is stupid
Text1.SaveFile "C:\fakevm.tmp", rtfText

'now loop through the lines
For i = 1 To Lines
    LineData = ReadLine("c:\fakevm.tmp", i)
  '  Debug.Print i & ": " & LineData
    Select Case Left(LineData, 3)
        Case "lda"  'load accumaltor from address
            'Need to check for the memory location now
            'Sorry, only direct addressing at the moment!
            MemoryTemp = Right(LineData, 1)
            'now we need to set the rom
            Rom(RomCounter).Data = lda
            Rom(RomCounter + 1).Data = MemoryTemp
            'incremement so we write to correct ROM next loop
            RomCounter = RomCounter + 2
        Case "ldx"  'load x from address
            'Need to check for the memory location now
            'Sorry, only direct addressing at the moment!
            MemoryTemp = Right(LineData, 1)
            'now we need to set the rom
            Rom(RomCounter).Data = ldx
            Rom(RomCounter + 1).Data = MemoryTemp
            RomCounter = RomCounter + 2
        Case "sta"  'store from res to address
            'Need to check for the memory location now
            'Sorry, only direct addressing at the moment!
            MemoryTemp = Right(LineData, 1)
            'now we need to set the rom
            Rom(RomCounter).Data = sta
            Rom(RomCounter + 1).Data = MemoryTemp
                RomCounter = RomCounter + 2
        Case "adx"  'add x to acc
            'set the rom value
            Rom(RomCounter).Data = adx
            RomCounter = RomCounter + 1
        Case "sbx"  'sub x from acc
            'set the rom value
            Rom(RomCounter).Data = sbx
            RomCounter = RomCounter + 1
        Case "hlt"  'Achtung! Dieeee!!!
            'set the rom value
            Rom(RomCounter).Data = hlt
            RomCounter = RomCounter + 1
    End Select

Next i

'Delete the temp file as we dont need it anymore
Kill "c:\fakevm.tmp"

''now display the address data from the rom in debug
'For r = 0 To 15
'    Debug.Print "Rom " & r & ": " & Rom(r).Data
'Next r

Debug.Print "Used " & RomCounter & " bytes of ROM."


'Woohoo, we did it.
'Our programme is in the ROM now, waiting for the control unit to deal with it.
'This is where the voodoo magic happens - I am going to fluff it in VB, i.e. VB
'will pretend to be the control unit and pretend to be the circuitry and the
'registers and the ALU etc, and magically come up with the results. How special.

'Enable the execute button so we can do it.
Command2.Enabled = True
'enable the view rom button so they can see it.
Command4.Enabled = True
End Sub

Public Function ReadLine(fName As String, LineNumber As Integer) As String
Dim oFSO As New FileSystemObject
Dim oFSTR As Scripting.TextStream
Dim ret As Long
Dim lCtr As Long

'Cant remember where this sub came from.  Thanks to whoever did it! :-)

If oFSO.FileExists(fName) Then
   Set oFSTR = oFSO.OpenTextFile(fName)
   Do While Not oFSTR.AtEndOfStream
      lCtr = lCtr + 1
      If lCtr = LineNumber Then
          ReadLine = oFSTR.ReadLine
          Exit Do
      End If
      oFSTR.SkipLine
  Loop
   oFSTR.Close
   Set oFSTR = Nothing

 End If
End Function

Private Sub Command2_Click()
'This is where it will all happen.  If this was a real cpu, this would be
'the contorl unit and would control the transfers between registers and
'the memory etc.

Dim adxtemp1, adxtemp2
On Error GoTo e
'We are running!  Red alert!
Stopped = False

'Disable the RAM boxes to stop sticky fingers ruining things,
'whilst we copy the values into an internal array
For r = 0 To 5
    RamT(r).Enabled = False
    Ram(r).Data = RamT(r).Text
Next r


'Right, now the business loop!
'This will look at the value in the rom and decide what to do.
'e.g., if we look in the rom and we get a 2 we know that is an
'ldx command.  We also know that the ldx command has an address
'with it, which according to the specification of this machine
'(i.e. the one I made up for it), should be in the next rom
'address.  We get this address then look in the relavent ram
'address and store the number into the x register. Its a big and
'long winded routine (and ugly!) but it should be easy to understand...

'I'm not bothering with a software instruction register as its a
'pain in the arse :-)  A variable will do!

'Start at 0 because we have rom(0).  i will act as the program counter
'register, when we have read an address we jump the counter up a couple
'and it should get the next instruction and 'skip' the memory address
For i = 0 To 15
    
    Select Case Rom(i).Data
        Case "1"  'lda
            'ok, we need an address for the data here.
            'Find outwhat it is and act accordingly
            Select Case Rom(i + 1).Data
                Case "A"
                    Acc.Data = Ram(0).Data
                Case "B"
                    Acc.Data = Ram(1).Data
                Case "C"
                    Acc.Data = Ram(2).Data
                Case "D"
                    Acc.Data = Ram(3).Data
                Case "E"
                    Acc.Data = Ram(4).Data
                Case "F"
                    Acc.Data = Ram(5).Data
            End Select
            'increment
            i = i + 1
        Case "2"  'ldx
            'ok, we need an address for the data here
            Select Case Rom(i + 1).Data
                Case "A"
                    XReg.Data = Ram(0).Data
                Case "B"
                    XReg.Data = Ram(1).Data
                Case "C"
                    XReg.Data = Ram(2).Data
                Case "D"
                    XReg.Data = Ram(3).Data
                Case "E"
                    XReg.Data = Ram(4).Data
                Case "F"
                    XReg.Data = Ram(5).Data
            End Select
            'increment
            i = i + 1
        Case "3"    'adx
            'this is just add, its easy
            'however, we have to convert the values in the
            'registers to hex before we do try to add
            adxtemp1 = "&h" + Acc.Data
            adxtemp2 = "&h" + XReg.Data
            Res.Data = Int(adxtemp1) + Int(adxtemp2)
            'no increment because we only used 1 rom byte on adx
        Case "4"    'sbx
            'this is just subtract, its easy
            'however, we have to convert the values in the
            'registers to hex before we do try to subtract
            adxtemp1 = "&h" + Acc.Data
            adxtemp2 = "&h" + XReg.Data
            Res.Data = Int(adxtemp1) - Int(adxtemp2)
            'no increment because we only used 1 rom byte on sbx
        Case "5"  'sta
            'ok, we need an address for the data here
            Select Case Rom(i + 1).Data
                Case "A"
                    Ram(0).Data = Res.Data
                Case "B"
                    Ram(1).Data = Res.Data
                Case "C"
                    Ram(2).Data = Res.Data
                Case "D"
                    Ram(3).Data = Res.Data
                Case "E"
                    Ram(4).Data = Res.Data
                Case "F"
                    Ram(5).Data = Res.Data
            End Select
            'increment
            i = i + 1
        Case "0" 'hlt
            'ok we got a hlt.
            'Thats all the programmer wants us to do, so
            'we will jump out of the loop now because all
            'the other rom bytes that are left are also
            'hlt commands.
            GoTo done
    End Select
Next i
done:

'almost there, we now need to update the ram texts with the new values
'so we can see if it has worked or not.
For r = 0 To 5
    RamT(r).Enabled = True
    RamT(r).Text = Ram(r).Data
Next r
'Ok, we are done.
Stopped = True
Exit Sub
e:
'erk, error.  Probably a bad value in ram
'tell the user and stop running
Stopped = True
For r = 0 To 5
    RamT(r).Enabled = True
Next r
MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbInformation, "Execution Error!"
End Sub

Private Sub Command3_Click()
'list the opcodes for people to play with

Dim Msg As String

Msg = "OpCodes available:" & vbCrLf & _
    vbTab & "hlt - stop exectuing" & vbCrLf & _
    vbTab & "lda n - load to acc from n" & vbCrLf & _
    vbTab & "ldx n - load to x from n" & vbCrLf & _
    vbTab & "adx - add x to acc" & vbCrLf & _
    vbTab & "sbx - subtract x from acc" & vbCrLf & _
    vbTab & "sta n - load to n from res" & vbCrLf & _
"Where n is an address in the RAM (in hex, e.g. A) using direct addressing."
MsgBox Msg, vbInformation, "OpCodes"

End Sub

Private Sub Command4_Click()
'They want to see the rom!
'Easy ;-)

Dim Msg As Variant
'loop through and add all of the data
For r = 0 To 15
    Msg = Msg + Rom(r).Data
Next r
'show it.
MsgBox "Contents of the ROM:" & vbCrLf & "(i.e., the programme)" & vbCrLf & vbCrLf & Msg, vbInformation, "ROM Contents"
End Sub

Private Sub Form_Load()
Stopped = True
'load in default programme
With Text1
    .SelText = "lda A" & vbCrLf
    .SelText = "ldx B" & vbCrLf
    .SelText = "adx" & vbCrLf
    .SelText = "sta F" & vbCrLf
    .SelText = "hlt"
End With

'make all ROM entries 0 (i.e. hlt)
For r = 0 To 11
    Rom(r).Data = "0"
Next r


End Sub


