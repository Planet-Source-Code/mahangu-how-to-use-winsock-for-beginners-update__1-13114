<div align="center">

## How To Use Winsock for Beginners \(Update\)


</div>

### Description

This tutorial shows newbies to VB, basically everything they need to know about Winsock. It shows how to open and close a Winsock connection and also how to send and receive data via a Winsock connection. Very easy to understand and use. Includes sample code that can be used in your applications! A few bugs have been fixed in this update so try it out!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-11-26 04:39:52
**By**             |[Mahangu](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mahangu.md)
**Level**          |Beginner
**User Rating**    |4.2 (38 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1206111252000\.zip](https://github.com/Planet-Source-Code/mahangu-how-to-use-winsock-for-beginners-update__1-13114/archive/master.zip)





### Source Code

<p align="center"><b><font face="Arial" color="#000080" size="5">Winsock for
Beginners</font></b></p>
<p align="center">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Introduction</b></font></p>
<p align="left"><font color="#000000" face="Arial">This tutorial will show
newcomers to Visual Basic how to use the Winsock ActiveX Control to transfer
data across the internet. This tutorial show beginners how to start a Winsock
connection, how to send data across a Winsock connection, how to receive data
using a Winsock Connection and how to close a Winsock connection.</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Why I wrote this tutorial</b></font></p>
<p align="left"><font face="Arial" color="#000000">I got asked a few questions
on Winsock so I decided to write a tutorial that would describe the very basics
of using Winsock. Also I thought that it would help new coders who were trying
to send data over the net.</font></p>
<p align="left">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="center"><b><font face="Arial" color="#000080" size="4">Getting Started</font></b></p>
<p align="left"><font face="Arial" color="#000000">1)Start VB and choose
'Standard EXE'</font></p>
<p align="left"><font face="Arial" color="#000000">2)Now Using the Add
Components (Right Click on Toolbar) add the Microsoft Winsock Control</font></p>
<p align="left"><font face="Arial" color="#000000">3)Double Click the New Icon
that Appeared on the Toolbar</font></p>
<p align="left"><font face="Arial" color="#000000">Now you will see the control
on the form. You can rename the control but in the code I will call it Winsock1.&nbsp;</font></p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<p align="center"><b><font face="Arial" color="#000080" size="4">Opening a
Winsock Connection</font></b></p>
<p align="left"><font face="Arial">To Open a Winsock Connection all you need to
do is to type Winsock1.Connect . But there are two values you have to give for
the code to work. Remote Host and Remote Port.</font></p>
<p align="left"><font face="Arial" color="#000000">Paste this Into the Form_Load()
, Command1_Click() or any other Sub</font></p>
<p align="left"><font face="Arial" color="#000000">'&lt;---- The Code Starts
Here ----&gt;</font></p>
<p align="left"><font face="Arial" color="#000000"><i>Winsock1.Connect , RemHost,
RemotePort,</i></font></p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Ends Here
----&gt;</font></p>
<p align="left"><font face="Arial" color="#000000">RemHost stands for the Remote
Host you want to connect to. The RemotePort stands for the Remote Port you want
to connect to.</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080" size="2"><b>Example</b></font></p>
<p align="left"><i><font face="Arial" color="#000000">Winsock1.Connect , &quot;127.0.0.1&quot; ,
&quot;100&quot; </font><font face="Arial" color="#008000">'This code example will
connect you to your own computer on Port 100&nbsp;</font></i><font size="1" face="Arial" color="#008000"><b>&nbsp;&nbsp;&nbsp;&nbsp;</b></font></p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<p align="center"><font face="Arial" color="#000080" size="4">Sending Data Using
Winsock</font></p>
<p align="left"><font face="Arial" color="#000000">Sending data using Winsock is
also relatively simple. Just use Winsock1.SendData . But this too requires a
value to be given. In plain English - It has to to know what data to send.</font></p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Starts Here
----&gt;</font></p>
<p align="left"><font face="Arial" color="#000000"><i>Winsock1.SendData(Data)</i></font></p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Ends Here
----&gt;</font></p>
<p align="left"><font face="Arial" color="#000000">Data stands for the data you
want to send.</font></p>
<p align="left"><font face="Arial" color="#000080" size="2"><b>Example</b></font></p>
<p align="left"><i><font face="Arial" color="#000000">Winsock1.SendData(&quot;Test&quot;)
</font><font face="Arial" color="#008000">'This code will send the data string
&quot;Test&quot;</font></i></p>
<p align="left">&nbsp;</p>
<p align="center"><font face="Arial" color="#000080" size="4">Receiving Data
Using Winsock&nbsp;</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000000">Receiving data using Winsock
is relatively more complex than the methods mentioned above. It requires code in
three places.&nbsp; It requires code in the Form_Load (or any other section), code in the Winsock1_DataArrival Section
, and code in the Winsock_ConnectionRequest event.&nbsp;</font></p>
<p align="left"><font face="Arial" color="#000080" size="3"><b>Step1 (Placing
the code in Form_Load event)</b></font></p>
<p align="left"><font face="Arial" color="#000000">Placing this code depends on when you want to start
accepting data. The best place to put this code is usually in the Form_Load
event.</font></p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Starts Here
----&gt;</font></p>
<p align="left"><font face="Arial" color="#000000"><i>Winsock1.LocalPort =
PortNumber</i></font></p>
<p align="left"><font face="Arial" color="#000000"><i>Winsock.Listen</i></font></p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Ends Here
----&gt;</font></p>
<p align="left"><font face="Arial" color="#000000">Data stands for the data you
want to send.</font></p>
<p align="left">&nbsp;</p>
<p align="left"><b><font face="Arial" color="#000080" size="2">Example</font></b></p>
<p align="left"><i><font face="Arial" color="#000000">Winsock1.LocalPort = 1000 </font><font face="Arial" color="#008000">'This
will set the port number to 1000</font></i></p>
<p align="left"><i><font face="Arial">Winsock.Listen '<font color="#008000">This
will tell Winsock to start listening</font></font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080" size="3"><b>Step 2 (Placing
the code in Winsock1_DataArrival Section)</b></font></p>
<p align="left"><font face="Arial" size="3" color="#000000">You will need to
place some code in the Winsock1_DataArrival event to tell Winsock what to do
once it receives data.</font></p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Starts Here
----&gt;</font></p>
<p align="left"><font face="Arial" color="#000000"><i>Winsock1.GetData (data)</i></font></p>
<p align="left"><i><font face="Arial" color="#000000">&nbsp;MsgBox&nbsp; (data) </font><font face="Arial" color="#008000">'This
will show the data in a Message Box</font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Ends Here
----&gt;</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" size="2" color="#000080"><b>Example</b></font></p>
<p align="left"><font face="Arial"><i>Dim StrData <font color="#008000">'This
declares the data string (can be place in general declarations too)</font></i></font></p>
<p align="left"><i><font face="Arial" color="#000000">Winsock1.GetData StrData </font><font face="Arial" color="#008000">'Tells
Winsock to get the data from the Port and put it in the data string</font></i></p>
<p align="left"><i><font face="Arial" color="#000000">&nbsp;MsgBox&nbsp; SrtData
</font><font face="Arial" color="#008000">'Displays the data in a Message Box</font></i></p>
<p align="center">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080" size="3"><b>Step 3 (Placing
the code in Winsock1_Connection Request Section)</b></font></p>
<p align="left"><font face="Arial" size="3" color="#000000">You will need to
place some code in the Winsock1_ConnectionRequest event to tell Winsock what do
when it receives a connection request.</font></p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Starts Here
----&gt;</font></p>
<p align="left"><i><font face="Arial">Dim RequestID <font color="#008000">'Declare
the RequestID String</font></font></i></p>
<p align="left"><font face="Arial"><i>If socket.State &lt;> sckClosed Then&nbsp;<br>
socket.Close<br>
socket.Accept requestID<br>
End If<br>
</i></font></p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Ends Here
----&gt;</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" size="2" color="#000080"><b>Example</b></font></p>
<p align="left"><i><font face="Arial">Dim RequestID <font color="#008000">Declare
the RequestID String</font></font></i></p>
<p align="left"><i><font face="Arial">If socket.State &lt;> sckClosed Then <font color="#008000">'If
Winsock is not closed</font><br>
socket.Close '<font color="#008000">Then Close the Connetion</font><br>
socket.Accept requestID&nbsp; <font color="#008000">Reuquest the ID&nbsp;</font><br>
End If<br>
</font></i></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" color="#000080" size="4">Closing a Winsock
Connection</font></p>
<p align="center"><font face="Arial">This is relatively simple. All you have to
do is to type one line of code. This can be place in almost any event on the
form including Form_Unload , Comman1_Click and so on.</font></p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Starts Here
----&gt;</font></p>
<p align="left"><i><font face="Arial" color="#000000">Winsock1.Close </font><font face="Arial" color="#008000">'Closes
the Winsock Connection</font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000000">&lt;---- The Code Ends Here
----&gt;</font></p>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" color="#000080" size="4">The End</font></p>
<p align="center"><font face="Arial" color="#000000">Please tell me how I can
improve this tutorial. If you have any questions or comments please post them
here and I will reply to them as soon as I can.</font></p>
</body>

