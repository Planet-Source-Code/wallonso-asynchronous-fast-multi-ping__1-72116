asynchronous Ping multiple Hosts using the IcmpSendEcho2 function
It looks a bit like multithreading.
And it is impressive, how less time is needed to ping up to 100 hosts.


Why IcmpSendEcho2 and not the standard IcmpSendEcho ?
If you use IcmpSendEcho your computer seems to be locked, while trying to get a ping response.
Because IcmpSendEcho is working synchronous, this means that the Caller has to wait for the Result.

IcmpSendEcho2 works asynchronous. So you can start to ping an do something else.

This is a tribute to LiTe's fast ParaPing(the basic idea).
Except I use ICMP instead of a specific port
and this is no real multithreading !

Comments are welcome