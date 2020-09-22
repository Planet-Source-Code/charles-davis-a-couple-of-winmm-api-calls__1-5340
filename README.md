<div align="center">

## A Couple of Winmm API Calls


</div>

### Description

There are 2 API Calls to the Winnmm API. One Detects if a Sound Card is installed. The other Plays an .AVI. You need to have Windows Media Player installed.
 
### More Info
 
This code is very simple and pretty self explanatory.

As far as I know there are none.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Charles Davis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/charles-davis.md)
**Level**          |Intermediate
**User Rating**    |4.0 (44 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/charles-davis-a-couple-of-winmm-api-calls__1-5340/archive/master.zip)

### API Declarations

```
'For SoundCard Function
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
'To Play Avi
Private Declare Function mciSendString Lib "winmm.dll" Alias _
"mciSendStringA" (ByVal lpstrCommand As String, _
ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long
```


### Source Code

```
Public Function SoundCard() As Boolean
Dim lng As Long
 lng = waveOutGetNumDevs()
 If lng > 0 Then
  SoundCard = True
  Exit Function
 Else
   SoundCard = False
   Exit Function
 End If
End Function
Public Sub PlayAvi()
Dim strAviPath As String
Dim strCmdStr As String
Dim lngReturnVal As Long
 strAviPath = "C:\winnt\clock.avi"
 strCmdStr = "play " & strAviPath & " fullscreen "
 lngReturnVal = mciSendString(strCmdStr, 0&, 0, 0&)
End Sub
```

