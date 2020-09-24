<div align="center">

## Internet File Transfer From Website


</div>

### Description

This is a module for easily adding the ability to transfer a file from a webserver to a hard drive. Good for a "live update" type of functionality or for downloading banners to your program. By Erick Jones, http://www.webdataconsultants.com
 
### More Info
 
You must provide a URL of the source file and a filename for the downloaded file name

First of all, this is a module. You'll have to add the module to your project. You also have to add a Microsoft Internet Transfer Control (and a Label1 Control if you are using the example below) to your form.

EXAMPLE USAGE:

Label1.Caption = "Getting File..."

TransferFile = Xfer(Inet1, "http://www.webdata.addr.com/index.html", App.Path + "\index.html")

Label1.Caption = InetStatus

InetStatus returns the status of the file transfer or an error if the transfer is unsuccessful.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[webdataconsultants](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/webdataconsultants.md)
**Level**          |Advanced
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/webdataconsultants-internet-file-transfer-from-website__1-43476/archive/master.zip)

### API Declarations

```
Global TransferFile As Boolean
Global b() As Byte
Global InetStatus As String
```


### Source Code

```
'####################################
'# Created 2002 Webdata Consultants #
'# You are free to use this module #
'# as long as you keep this header #
'# intact.       #
'# www.webdataconsultants.com  #
'####################################
'
'First of all, to use this module you have
'add a Microsoft Internet Transfer Control
'and a Label1 Control to your form.
'
'EXAMPLE USAGE:
 'Label1.Caption = "Getting File..."
 'TransferFile = Xfer(Inet1, "http://www.webdata.addr.com/index.html", App.Path + "\index.html")
 'Label1.Caption = InetStatus
Global TransferFile As Boolean
Global b() As Byte
Global InetStatus As String
Public Function Xfer(Inet1 As Inet, strURL As String, InputFile As String) As Boolean
 On Error GoTo ErrorHandle
 Inet1.AccessType = icUseDefault
 Inet1.RequestTimeout = 10 'Higher number increases request timeout
 b() = Inet1.OpenURL(strURL, icByteArray)
 Open InputFile For Binary Access _
 Write As #1
 Put #1, , b()
 Close #1
 InetStatus = "Done"
Exit Function
ErrorHandle:
 Select Case Err.Number
  Case 75
   InetStatus = "Destination file is read-only."
  Case 35761
   InetStatus = "Request timed out. Please check your internet connection or try again later."
  Case Else
   InetStatus = "Error: " & Err.Number
 End Select
End Function
```

