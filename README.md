<div align="center">

## Using NetFileEnum to determine which files are open on server


</div>

### Description

Find out who is currently using which file on the server. I am using Windows NETFILEENUM CALL. The structure that i have used in this example is only for WIN NT/2000. If your are going to use WIN95/WIN98/ME grab the structure from msdn file_info_50. Please vote!!!!!
 
### More Info
 
Computername


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Khursheed\_Siddiqui](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/khursheed-siddiqui.md)
**Level**          |Intermediate
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/khursheed-siddiqui-using-netfileenum-to-determine-which-files-are-open-on-server__1-37291/archive/master.zip)

### API Declarations

```
Option Explicit
Private Const NERR_SUCCESS As Long = 0&
Private Const MAX_PREFERRED_LENGTH As Long = -1
Private Const ERROR_MORE_DATA As Long = 234&
Private Type FILE_INFO_3
fi3_id As Long
fi3_permissions As Long
fi3_num_locks As Long
fi3_Pathname As Long
fi3_username As Long
End Type
Private Declare Function NetFileEnum Lib "netapi32" _
(ByVal Servername As Long, _
ByVal Basepath As Long, _
ByVal Username As Long, _
ByVal level As Long, _
bufptr As Long, _
ByVal prefmaxlen As Long, _
entriesread As Long, _
totalentries As Long, _
resume_handle As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" _
(ByVal Buffer As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" _
(pTo As Any, uFrom As Any, _
ByVal lSize As Long)
Private Declare Function lstrlenW Lib "kernel32" _
(ByVal lpString As Long) As Long
```


### Source Code

```
Private Function GetPointerToByteStringW(ByVal dwData As Long) As String
 Dim tmp() As Byte
 Dim tmplen As Long
 If dwData <> 0 Then
 tmplen = lstrlenW(dwData) * 2
 If tmplen <> 0 Then
  ReDim tmp(0 To (tmplen - 1)) As Byte
  CopyMemory tmp(0), ByVal dwData, tmplen
  GetPointerToByteStringW = tmp
 End If
 End If
End Function
Private Sub Form_Load()
 Dim bufptr As Long
 Dim dwServer As Long
 Dim dwEntriesread As Long
 Dim dwTotalentries As Long
 Dim dwResumehandle As Long
 Dim success As Long
 Dim nStructSize As Long
 Dim cnt As Long
 Dim usrname As String
 Dim bServer As String
 Dim fi3 As FILE_INFO_3
 bServer = "\\" & ComputerName & vbNullString
 dwServer = StrPtr(bServer)
 success = NetFileEnum(dwServer, 0&, 0&, 3,bufptr,MAX_PREFERRED_LENGTH, dwEntriesread,dwTotalentries,dwResumehandle)
 If success = NERR_SUCCESS And success <> ERROR_MORE_DATA Then
 nStructSize = LenB(fi3)
 For cnt = 0 To dwEntriesread - 1
  CopyMemory fi3, ByVal bufptr + (nStructSize * cnt), nStructSize
  usrname = GetPointerToByteStringW(fi3.fi3_username)
  If Len(usrname) > 0 Then
  Print usrname & vbTab & fi3.fi3_permissions & vbTab & fi3.fi3_num_locks & vbTab & GetPointerToByteStringW(fi3.fi3_Pathname)
  End If
 Next
 End If
Call NetApiBufferFree(bufptr)
End Sub
```

