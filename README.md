<div align="center">

## Read Write


</div>

### Description

read a text file and convert it into an array
 
### More Info
 
csv txt

array


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brent Luyet](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brent-luyet.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brent-luyet-read-write__1-45794/archive/master.zip)

### API Declarations

micorosoft scripting runtime


### Source Code

```

'Purpose:This class is used to read and create files text files
'Dependancies:must add Reference Microsoft Scripting Runtime
'Creation Date: ?
'Author: Brent Luyet
'Revision  Date  Revision By
'1.02    3/30/03 BJL
Public Function read(InFile As String, outArray() As String)
'Purpose:This function receives a filename and returns values in an array
'Creation Date:?
'Author: Brent Luyet
'Revision  Date  Revision By
'1.01    1/30/03 mst
Dim fso As New FileSystemObject   'must add Reference Microsoft Scripting Runtime
Dim fts As TextStream
Dim inString As String       'temp value to hold string read from file
Set fts = fso.OpenTextFile(InFile, ForReading, False) 'open infile for read only
inString = fts.ReadAll       'read entire file into inString
outArray = Split(inString, vbCrLf) 'split inString into outArray
'Clean up
fts.Close
Set fts = Nothing
Set fso = Nothing
End Function
Public Function WriteFile(ByVal OutFile As String, ByRef outArray() As String)
'Purpose:This function will create a file using the values passed in the value out array
'Creation Date: Date Class was created
'Author: Brent Lyute
'Revision  Date  Revision By
'1.01    1/30/03 mst
Dim fso As New FileSystemObject 'must add Reference Microsoft Scripting Runtime
Dim fts As TextStream
Dim OutString As String     'temp val for holding array before writing to file
OutString = Join(WriteOut, vbCrLf) 'join array to temp string outString
Set fts = fso.OpenTextFile(OutFile, ForWriting, True) 'Open OutFile for write, overwrite
fts.Write OutString ' Write temp string outString to OutFile
'Clean UP
fts.Close
Set fts = Nothing
Set fso = Nothing
End Function
```

