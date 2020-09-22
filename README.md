<div align="center">

## Simple Winsock Tutorial


</div>

### Description

This is a short and to the point winsock tutorial. Pleae note this only covers tranferring text between two computers. Nothing more.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Codeaholic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/codeaholic.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/codeaholic-simple-winsock-tutorial__1-35422/archive/master.zip)





### Source Code

<html>
<head>
<title>Winsock Tutorial</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p> </p>
<p><font size="4"><b>WINSOCK TUTORIAL</b></font></p>
<p> </p>
<p><font face="Arial, Helvetica, sans-serif">Ok, heres how to use the Winsock
 control included with Visual Basic. </font></p>
<p><font face="Arial, Helvetica, sans-serif">First add a Winsock control to a
 blank form, and then add two command buttons to the form. <br>
 Double click one of the command buttons and insert the following code:<br>
 </font></p>
<p><font face="Arial, Helvetica, sans-serif">Private Sub Command1_Click()<br>
 Winsock1.Connect "localhost", 3000<br>
 End Sub<br>
 </font></p>
<p><font face="Arial, Helvetica, sans-serif">This code simply tells the Winsock
 control to connect to the IP of your computer (if you wanted to connect to another
 computer you would substitute localhost for the IP of the computer you wanted
 to connect to). This code also sets the port to connect on. Something above
 3000 is good because below this there is alot of service port used by windows
 eg. netbios<br>
 Now double click the other command button and add the following code:<br>
 </font></p>
<p><font face="Arial, Helvetica, sans-serif">Private Sub Command2_Click()<br>
 Winsock1.SendData "hello this is a test"<br>
 End Sub<br>
 <br>
 This bit of code is self-explanatory. It sends whatever you tell it.<br>
 Right.... we've now coded the client program. Now for the server. Minimize the
 project you've just created and start a new one. <br>
 Double click on the form and choose the Form_Activate event. Insert the code
 like this:<br>
 </font></p>
<p><font face="Arial, Helvetica, sans-serif">Private Sub Form_Activate()<br>
 Winsock1.LocalPort = 3000<br>
 Winsock1.Listen<br>
 End Sub<br>
 </font></p>
<p><font face="Arial, Helvetica, sans-serif">This code defines the port to listen
 on and sets the Winsock controls state to listen.<br>
 Now insert this code:<br>
 </font></p>
<p><font face="Arial, Helvetica, sans-serif">Private Sub Winsock1_ConnectionRequest(ByVal
 requestID As Long)<br>
 Winsock1.Close<br>
 Winsock1.Accept requestID<br>
 End Sub<br>
 </font></p>
<p><font face="Arial, Helvetica, sans-serif">This code accepts a connection.<br>
 Add this code as well: <br>
 </font></p>
<p><font face="Arial, Helvetica, sans-serif">Private Sub Winsock1_DataArrival(ByVal
 bytesTotal As Long)<br>
 Dim n As String<br>
 Winsock1.GetData n<br>
 MsgBox n<br>
 End Sub<br>
 </font></p>
<p><font face="Arial, Helvetica, sans-serif">This code gets the data received
 from the other program when you click command2. Please not you <b>must declare
 the variable used to store the data you got. </b> If you do not you will just
 get ? Marks. </font></p>
<p><font face="Arial, Helvetica, sans-serif">Now run both of the programs and
 click the first command button to connect and the second one to send data to
 the other program. The other program should show the message you told the first
 program to send.</font></p>
<p> </p>
<p></p>
<p><font face="Arial, Helvetica, sans-serif"><br>
 </font> </p>
</body>
</html>

