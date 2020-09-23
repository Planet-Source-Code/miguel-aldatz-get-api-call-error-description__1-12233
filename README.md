<div align="center">

## Get api call error description


</div>

### Description

when call Api funtion, this call return a long value, this value indicate the error code of the function. This code get description for this code error, in the windows standart error descriptions.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-10-22 17:29:40
**By**             |[Miguel Aldatz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/miguel-aldatz.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1087710222000\.zip](https://github.com/Planet-Source-Code/miguel-aldatz-get-api-call-error-description__1-12233/archive/master.zip)

### API Declarations

```
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Any) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
```





