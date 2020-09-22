<div align="center">

## Fastest FileLoading Techniquie for simple VB


</div>

### Description

In All the other methods I've tried, I run into a barrier(around 170K+) where loading a file suddenly becomes very slow. I've tried many methods found here and in books, and none have proven to work as well as I was looking for.

After a long search I ended up spending some off-time figuring a way out myself. This is what I came up with.(speed increases from 20 seconds@170k to .4 seconds@170k)
 
### More Info
 
Parameter 1: Path

Path is the path to the file. If there is no file, the program will return nothing.

Remeber that strings in VB are limited. I believe it is somwhere around 2.1gigs. if you happen to be authoring an application for a machine that can handle that much, you should seriously consider an alternate method.

It returns a binary string(an ordinary string, just that the bytes contain any character, not just Ansi Text). the file is completely loaded into memory.

Remember that this is probably the fastest way to load a file into memory. if your file is larger than the available memory, then it will become slow, and a major burden on the system.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Violet Sagmiller](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-violet-sagmiller.md)
**Level**          |Intermediate
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-violet-sagmiller-fastest-fileloading-techniquie-for-simple-vb__1-23401/archive/master.zip)





### Source Code

```
Public Function LoadBin(Path As String) As String
On Error GoTo hell' isn't that where errors belong?
Dim nfile As String ' This becomes the file memory
Dim i As Long ' temp int
i = FreeFile ' Gets a free file number so that this code doesn't interfere with anything else.
Open Path For Binary As i ' read the file raw
  nfile = String(LOF(i), " ") ' create a string in memory that is the size of the file.
  Get i, , nfile ' in one pass, load the entire file as a single record.
Close i ' clean up the mess
LoadBin = nfile 'set the return value
hell: ' this is where it goes if the code breaks anyway.
End Function
```

