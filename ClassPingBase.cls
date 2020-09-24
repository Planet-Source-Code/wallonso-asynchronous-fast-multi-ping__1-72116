VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPingBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type ip_option_information
    ttl             As Byte 'Byte     'Time To Live
    Tos             As Byte     'Type Of Service
    Flags           As Byte     'IP header flags
    OptionsSize     As Byte     'Size in bytes of options data
    OptionsData     As Long     'Pointer to options data
End Type

Private Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    datasize        As Long 'formerly integer
   'Reserved        As Integer
    DataPointer     As Long
    Options         As ip_option_information
    Data            As String * 250
End Type

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function inet_addr Lib "wsock32" (ByVal s As String) As Long

Private Declare Function IcmpSendEcho2 Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal hEvent As Long, ByVal ApcRoutine As Any, ByVal ApcContext As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ByVal ReplyBuffer As Long, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long

Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
'Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
'Private Declare Function WaitForMultipleObjectsArray Lib "kernel32" Alias "WaitForMultipleObjects" (ByVal nCount As Long, lpHandles() As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Private Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private Const WAIT_OBJECT_0 As Long = 0
Private Const WAIT_TIMEOUT As Long = &H102
Private Const WAIT_ABANDONED = &H80
Private Const WAIT_FAILED = -1&
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Event PingSuccess(sIPAdress As String, lNewStatus As Long, ArrayIndex As Long)
Public Event PingFail(sIPAdress As String, lNewStatus As Long, ArrayIndex As Long)

Private mvarNumParalellActions As Long
Private mvarListPingRunning As Boolean

Public Property Let NumParalellActions(ByVal NewData As Long)
    mvarNumParalellActions = NewData
End Property
Public Property Get NumParalellActions() As Long
    NumParalellActions = mvarNumParalellActions
End Property

Public Property Get ListPingRunning() As Boolean
    ListPingRunning = mvarListPingRunning
End Property

Public Function PingHostSingle(sHostIP As String, TimeOut As Long) As Long
Const sSendData As String = "TESTMESSAGE"
Dim Buffer As ICMP_ECHO_REPLY
Dim lhwndPort As Long
Dim hEvent As Long
Dim sd As SECURITY_ATTRIBUTES
Dim WaitRes As Long
Dim tm1 As Long, tm2 As Long

    '//needed for Create Event
    With sd
        .nLength = Len(sd) 'we pass the length of sd
        .lpSecurityDescriptor = 0
        .bInheritHandle = 0
    End With

    '//We create an Event-ID to be registered with IcmpSendEcho2 and WaitForMultipleObjects to wait for
    hEvent = CreateEvent(sd, True, False, "PINGV2")

    lhwndPort = IcmpCreateFile()    '//Standard for ICMP, create a handle
    
    Call IcmpSendEcho2(lhwndPort, hEvent, Nothing, 0, inet_addr(sHostIP), sSendData, Len(sSendData), 0, VarPtr(Buffer), Len(Buffer), 1000)

    '//Gere we start to wait
    tm1 = GetTickCount()
    Do
        WaitRes = WaitForSingleObject(hEvent, 5) ' WaitForMultipleObjects(1, hEvent, 1, 10)
        tm2 = GetTickCount
        DoEvents
        If (tm2 - tm1) < 1000 Then
            Select Case WaitRes
                Case WAIT_ABANDONED, WAIT_FAILED
                    DoEvents
                    PingHostSingle = -2
                    Exit Do
                Case WAIT_OBJECT_0
                    If Buffer.Status = 0 Then
                        PingHostSingle = Buffer.RoundTripTime
                    Else
                        PingHostSingle = -1
                    End If
                    Exit Do
'//Cause were in a loop, this message will not work for us
'                Case WAIT_TIMEOUT
            End Select
            DoEvents
        Else
            DoEvents
            PingHostSingle = -1
            Exit Do
        End If
    Loop While (1)
    '//And clear all created things
    ResetEvent hEvent
    IcmpCloseHandle lhwndPort
End Function


Public Function PingHostList(saIpAdresses() As String, TimeOutHost As Long, TimeOutAll As Long) As Boolean
Dim NumAdresses As Long
Dim laHandlesICMP() As Long
Dim laEvents() As Long
Dim laTimerStart() As Long
Dim laTimerSop() As Long
Dim laWaitResults() As Long
Dim bUseResult() As Boolean
Dim baEchoSend() As Boolean
Dim aBuffers() As ICMP_ECHO_REPLY
Dim bProcessed() As Integer
Dim nProcessed As Long
Dim nLoop As Long
Dim baInitiated() As Boolean
Dim TimerStart As Long
Dim Timerstop As Long

Dim Buffer As ICMP_ECHO_REPLY
Dim sd As SECURITY_ATTRIBUTES
Const sSendData As String = "PINGDATAMESSAGE"
Dim WaitRes As Long

Dim nStart As Long, nEnd As Long, BoundCtr As Long

    '//Prevent from reentry
    If mvarListPingRunning = True Then
        PingHostList = False
        Exit Function
    End If
    mvarListPingRunning = True
    NumAdresses = UBound(saIpAdresses)
'//init the needed arrays
    ReDim laHandlesICMP(NumAdresses)
    ReDim laEvents(NumAdresses)
    ReDim laTimerStart(NumAdresses)
    ReDim laTimerSop(NumAdresses)
    ReDim laWaitResults(NumAdresses)
    ReDim baEchoSend(NumAdresses)
    ReDim aBuffers(NumAdresses)
    ReDim bProcessed(NumAdresses)
    
    '//needed for Create Event
    With sd
        .nLength = Len(sd) 'we pass the length of sd
        .lpSecurityDescriptor = 0
        .bInheritHandle = 0
    End With
    
    For nLoop = 0 To NumAdresses - 1
        '//We create an Event for WaitForMultipleObjects to wait for
        laEvents(nLoop) = CreateEvent(sd, True, False, "PINGV2C" & nLoop)
        laHandlesICMP(nLoop) = IcmpCreateFile()    '//Standard for ICMP
    Next
    TimerStart = GetTickCount
    nProcessed = 0  '//Reset Counter for processed hosts
    nStart = 0
    If NumAdresses > mvarNumParalellActions Then
        nEnd = mvarNumParalellActions
    Else
        nEnd = NumAdresses - 1
    End If
    BoundCtr = 0
    Do
        For nLoop = nStart To nEnd
            '//Did we already request an echo
            If baEchoSend(nLoop) = False Then
                baEchoSend(nLoop) = True    '//Lock this adress, so we don't create a second Ping
                Call IcmpSendEcho2(laHandlesICMP(nLoop), laEvents(nLoop), Nothing, 0, inet_addr(saIpAdresses(nLoop)), sSendData, Len(sSendData), 0, VarPtr(aBuffers(nLoop)), Len(aBuffers(nLoop)), TimeOutHost)
            End If
            If bProcessed(nLoop) = False Then
                laWaitResults(nLoop) = WaitForSingleObject(laEvents(nLoop), 5) '//We wait only for 5 Millisecs, to give all the other IPs Time to answer
                Select Case laWaitResults(nLoop)
                    Case WAIT_ABANDONED, WAIT_FAILED
                        DoEvents
                        bProcessed(nLoop) = -1
                        nProcessed = nProcessed + 1
                        BoundCtr = BoundCtr + 1
                    Case WAIT_OBJECT_0
                        bProcessed(nLoop) = 1
                        nProcessed = nProcessed + 1
                        BoundCtr = BoundCtr + 1
                        If aBuffers(nLoop).Status = 0 Then
                            '//Valid Info : So Host is reachable
                            RaiseEvent PingSuccess(saIpAdresses(nLoop), aBuffers(nLoop).RoundTripTime, nLoop)
                        Else
                            RaiseEvent PingFail(saIpAdresses(nLoop), -1, nLoop)
                        End If
    '//Cause were in a loop, this message will not work for us
'                    Case WAIT_TIMEOUT
                        'bProcessed(nLoop) = -1
                End Select
                End If
                DoEvents
            Next
        Timerstop = GetTickCount
        If (Timerstop - TimerStart) < TimeOutAll Then
            '
        Else
            DoEvents
            Debug.Print "TIMED OUT"
            Exit Do
        End If
        '//All done, no more Action needed
        If nProcessed >= NumAdresses - 1 Then
            Exit Do
        End If
        If nEnd < NumAdresses And BoundCtr >= mvarNumParalellActions Then
            nStart = nEnd + 1
            nEnd = nEnd + mvarNumParalellActions
            If nEnd > NumAdresses - 1 Then nEnd = NumAdresses - 1
            BoundCtr = 0
            Debug.Print "New bounds :"; nStart, nEnd, Time
        End If
    Loop While (1)
        
'//IMPORTANT : UNregister all messages/Evemnts and open ICMP-handles
    For nLoop = 0 To NumAdresses - 1
        ResetEvent laEvents(nLoop)
        IcmpCloseHandle (laHandlesICMP(nLoop))
        If bProcessed(nLoop) <= 0 Then
            '//notify caller with the rest of the host, that we haven't hit
            RaiseEvent PingFail(saIpAdresses(nLoop), -1, nLoop)
        End If
    Next
'//And we free up some memory
    Erase laHandlesICMP
    Erase laEvents
    Erase laTimerStart
    Erase laTimerSop
    Erase laWaitResults
    Erase baEchoSend
    Erase aBuffers
    Erase bProcessed
    mvarListPingRunning = False
End Function


Private Sub Class_Initialize()
    mvarNumParalellActions = 500
End Sub
