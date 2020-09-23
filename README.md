<div align="center">

## A Simple Encryption/Decryption bas using ASCII


</div>

### Description

This is great for beginners! This takes each character of a string and converts it to an ASCII value, then adds or substracts a designated number in order to decrypt/encrypt the string. Great for in-app passwords! Check it out

sVariable = ENCRYPT(sYourString,lLenOfString)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jaime Muscatelli](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jaime-muscatelli.md)
**Level**          |Beginner
**User Rating**    |4.6 (32 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jaime-muscatelli-a-simple-encryption-decryption-bas-using-ascii__1-42007/archive/master.zip)

### API Declarations

Option Explicit


### Source Code

```
Option Explicit
'*** By Jaime Muscatelli
' Just a simple example usage sub ( Main() is Not needed for prog!)
Public Sub Main()
Dim sString As String
Dim sEncrypted As String
Dim sDecrypted As String
sString = InputBox("String?", "String?")
sEncrypted = ENCRYPT(sString, Len(sString))
MsgBox sEncrypted
sDecrypted = DECRYPT(sEncrypted, Len(sEncrypted))
MsgBox sDecrypted
End Sub
' The encryption sub
' USAGE:  sVariable = ENCRYPT(sYourString, lLengthOfString)
Private Function ENCRYPT(sString As String, lLEn As Long) As String
'Just declaring variables
Dim I As Long
Dim NewChar As Long
I = 1 'can't start a string at 0 :-)
' Go through each character in the string and convert
' it to an ASCII value, add the number desired (here, 13), and
' then place it into a new string
Do Until I = lLEn + 1
NewChar = Asc(Mid(sString, I, 1)) + 13
ENCRYPT = ENCRYPT + Chr(NewChar)
I = I + 1
Loop
End Function
'Decryption sub
'USAGE: sVariable = DECRYPT(sYourstring, lLengthOfString)
Private Function DECRYPT(sString As String, lLEn As Long) As String
Dim I As Long
Dim NewChar As Long
I = 1
'Doing the reverse of the encryption sub!
Do Until I = lLEn + 1
NewChar = Asc(Mid(sString, I, 1)) - 13
DECRYPT = DECRYPT + Chr(NewChar)
I = I + 1
Loop
End Function
```

