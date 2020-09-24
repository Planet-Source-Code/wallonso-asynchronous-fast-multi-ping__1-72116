<div align="center">

## asynchronous Fast multi ping


</div>

### Description

asynchronous Ping of multiple Hosts using the IcmpSendEcho2 function

It looks a bit like multithreading.

And it is impressive, how less time is needed to ping up to 100 or more hosts.

(i tried 15 class C subnets (=over 3000 hosts), with 5 percent hosts online in 60 seconds)

Why IcmpSendEcho2 and not the standard IcmpSendEcho ?

If you use IcmpSendEcho your computer seems to be locked, while trying to get a ping response.

Because IcmpSendEcho is working synchronous, this means that the Caller has to wait for the Result.

IcmpSendEcho2 works asynchronous. So you can start to ping an do something else.

This is a tribute to LiTe's fast ParaPing(the basic idea).

Except I use ICMP instead of a specific port.

Comments are welcome
 
### More Info
 
see readme


<span>             |<span>
---                |---
**Submitted On**   |2009-05-25 20:34:02
**By**             |[Wallonso](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/wallonso.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Fast\_multi2153365272009\.zip](https://github.com/Planet-Source-Code/wallonso-asynchronous-fast-multi-ping__1-72116/archive/master.zip)

### API Declarations

see code





