<div align="center">

## Directory Cleaner \(recursively\)


</div>

### Description

This function attempts to delete all files

and subdirectories of the given

directory name, and leaves the given

directory intact, but completely empty.

If the Kill command generates an error (i.e.

file is in use by another process -

permission denied error), then that file and

subdirectory will be skipped, and the

program will continue (On Error Resume Next).

EXAMPLE CALL:

ClearDirectory "C:\Temp\"
 
### More Info
 
Full path directory name

Kill statement may error out for various reasons, which will prevent those files/directories from being deleted.

WARNING: If a subdirectory is prevented from being deleted (i.e. Kill statement errors out because of file access error), the loop will NOT terminate (While Len(sSubDir) > 0).

This may not be an issue in most cases, but I just wanted to make clear the limitations of this code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Gary Choma](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gary-choma.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gary-choma-directory-cleaner-recursively__1-25128/archive/master.zip)





### Source Code

```
Private Sub ClearDirectory(psDirName)
'This function attempts to delete all files
'and subdirectories of the given
'directory name, and leaves the given
'directory intact, but completely empty.
'
'If the Kill command generates an error (i.e.
'file is in use by another process -
'permission denied error), then that file and
'subdirectory will be skipped, and the
'program will continue (On Error Resume Next).
'
'EXAMPLE CALL:
' ClearDirectory "C:\Temp\"
Dim sSubDir
If Len(psDirName) > 0 Then
 If Right(psDirName, 1) <> "\" Then
 psDirName = psDirName & "\"
 End If
 'Attempt to remove any files in directory
 'with one command (if error, we'll
 'attempt to delete the files one at a
 'time later in the loop):
 On Error Resume Next
 Kill psDirName & "*.*"
 DoEvents
 sSubDir = Dir(psDirName, vbDirectory)
 Do While Len(sSubDir) > 0
 'Ignore the current directory and the
 'encompassing directory:
 If sSubDir <> "." And _
  sSubDir <> ".." Then
  'Use bitwise comparison to make
  'sure MyName is a directory:
  If (GetAttr(psDirName & sSubDir) And _
  vbDirectory) = vbDirectory Then
  'Use recursion to clear files
  'from subdir:
  ClearDirectory psDirName & _
   sSubDir & "\"
  'Remove directory once files
  'have been cleared (deleted)
  'from it:
  RmDir psDirName & sSubDir
  DoEvents
  'ReInitialize Dir Command
  'after using recursion:
  sSubDir = Dir(psDirName, vbDirectory)
  Else
  'This file is remaining because
  'most likely, the Kill statement
  'before this loop errored out
  'when attempting to delete all
  'the files at once in this
  'directory. This attempt to
  'delete a single file by itself
  'may work because another
  '(locked) file within this same
  'directory may have prevented
  '(non-locked) files from being
  'deleted:
  Kill psDirName & sSubDir
  sSubDir = Dir
  End If
 Else
  sSubDir = Dir
 End If
 Loop
End If
End Sub
```

