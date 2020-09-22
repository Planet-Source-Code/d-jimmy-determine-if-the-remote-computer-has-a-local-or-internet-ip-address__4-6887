<div align="center">

## Determine if the remote computer has a local or internet IP address


</div>

### Description

if you have to determine if the remote computer visitor is local or external of your site and according to the case, for exemple redirect the user to the correct page...
 
### More Info
 
This range of addresses shows constant IP of a LAN and never a computer will have this IP on internet (Private IP address ranges as specified in RFC 1918).

'10.0.0.0 -> 10.255.255.255

'172.16.0.0 -> 172.31.255.255

'192.168.0.0 -> 192.168.255.255

But don't forget :

A Lan administrator can configure a LAN computer with an address IP out of the LAN range above.

With disabling "Allow Anonymous Access" on your application in IIS, you can know if the user of the remote computer was logged (and be on the LAN).

To do that get in String variable the logon with request.serverVariables.item("LOGON_USER").

An empty String means the user is not logged.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[D\. Jimmy ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/d-jimmy.md)
**Level**          |Beginner
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/d-jimmy-determine-if-the-remote-computer-has-a-local-or-internet-ip-address__4-6887/archive/master.zip)





### Source Code

```
<HTML>
<BODY>
<%
dim ArrayIPLocalStart(2)
dim ArrayIPLocalEnd(2)
Dim ArrayIPClient
Dim IPClient
dim blnLocal
Ipclient = Request.ServerVariables("REMOTE_ADDR")
Response.Write "Your IP is " & ipclient & "<BR>"
blnLocal = false
' These IP address are constant IP of a LAN
ArrayIPLocalStart(0) = "010.000.000.000" '"10.0.0.0"
ArrayIPLocalEnd(0) = "010.255.255.255" '"10.255.255.255"
ArrayIPLocalStart(1) = "172.016.000.000" '"172.16.0.0"
ArrayIPLocalEnd(1) = "172.031.255.255" '"172.31.255.255"
ArrayIPLocalStart(2) = "192.168.000.000" '"192.168.0.0"
ArrayIPLocalEnd(2) = "192.168.255.255" '"192.168.255.255"
ArrayIPClient = split(Ipclient,".")
' format the remote IP like constant LAN IP
for i = lbound(ArrayIPClient) to ubound(ArrayIPClient)
	ArrayIPClient(i) = string(3-len(ArrayIPClient(i) ),"0") & ArrayIPClient(i)
next
IPClient = join(ArrayIPClient,"")
if trim(ipclient) <> "" then
	for i = lbound(ArrayIPLocalStart) to ubound(ArrayIPLocalStart)
		ArrayIPLocalStart(i) = replace(ArrayIPLocalStart(i),".","")
		ArrayIPLocalEnd(i) = replace(ArrayIPLocalEnd(i),".","")
		if IPClient >= ArrayIPLocalStart(i) and IPClient =< ArrayIPLocalEnd(i) then
			blnLocal = true
			exit for
		End if
	next
end if
if blnLocal then
	Response.Write " And your computer has a local IP"
else
	Response.Write " And your computer has a internet IP"
End if
%>
</BODY>
</HTML>
```

