<div align="center">

## A function that encrypts/decrypts\.


</div>

### Description

2 functions 1 that encrypts a passed variable(string) and the other decrypts the data(path of file)

Mainly used for a single password.

The code encrypts each character with a simple formula, the more characters the better the encryption.

The limit on the password length is 50 characters.
 
### More Info
 
The encrypt function input is the passed variable which is the password to encrypt.

Its very simple, and not even close to "great" encryption. but it gets the job done for what it does.

The decrypt function returns the decrypted value

No side effects


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[m@llot](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/m-llot.md)
**Level**          |Unknown
**User Rating**    |4.3 (30 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/m-llot-a-function-that-encrypts-decrypts__1-2616/archive/master.zip)





### Source Code

```
'simple just pass the password to it like this
'Encrypt("password")
Private Function Encrypt(varPass As String)
If Dir(path to save password to) <> "" Then: Kill "path to save password to"
Dim varEncrypt As String * 50
Dim varTmp As Double
 Open "path to save password to" For Random As #1 Len = 50
  For I = 1 To Len(varPass)
   varTmp = Asc(Mid$(varPass, I, 1))
   varEncrypt = Str$(((((varTmp * 1.5) / 2.1113) * 1.111119) * I))
   Put #1, I, varEncrypt
  Next I
 Close #1
End Function
'returns the decrypted pass
'like if decrypt() = "password" then
Private Function Decrypt()
Open "path to save password to" For Random As #1 Len = 50
  Dim varReturn As String * 50
  Dim varConvert As Double
  Dim varFinalPass As String
  Dim varKey As Integer
  For I = 1 To LOF(1) / 50
   Get #1, I, varReturn
   varConvert = Val(Trim(varReturn))
   varConvert = ((((varConvert / 1.5) * 2.1113) / 1.111119) / I)
   varFinalPass = varFinalPass & Chr(varConvert)
  Next I
  Decrypt = varFinalPass
 Close #1
End Function
```

