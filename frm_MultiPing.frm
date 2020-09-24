VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_MultiPing 
   Caption         =   "MultiPing"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ImageList ILState 
      Left            =   3960
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_MultiPing.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_MultiPing.frx":17D3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_MultiPing.frx":2FA74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_MultiPing.frx":362D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnPingAsync 
      Caption         =   "Ping Check"
      Height          =   735
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin MSComctlLib.ListView LV_IPs 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ILState"
      SmallIcons      =   "ILState"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton btnPingList 
      Caption         =   "Ping loop"
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton btnPing 
      Caption         =   "Ping"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "10.10.10.10"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer TmrPingS 
      Left            =   240
      Top             =   2520
   End
   Begin VB.Label lblHosts 
      Alignment       =   1  'Rechts
      Caption         =   "Hosts"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblHosts 
      Caption         =   "Hosts"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblPResult 
      Alignment       =   2  'Zentriert
      Caption         =   "X"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblIP 
      Caption         =   "IP-Address"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frm_MultiPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private WithEvents m_clsPingBase As ClassPingBase
Attribute m_clsPingBase.VB_VarHelpID = -1

Private m_bFirstRun As Boolean

Private Sub btnPing_Click()
Dim res As Long
    btnPing.Enabled = False
    res = m_clsPingBase.PingHostSingle(txtIP, 100)
    If res < 0 Then
        lblPResult.Caption = ""
        lblPResult.BackColor = vbRed
    Else
        lblPResult.Caption = res
        lblPResult.BackColor = vbGreen
        
    End If
    btnPing.Enabled = True
End Sub

Private Sub btnPingAsync_Click()
Dim saAdresses() As String, n As Integer
Dim tm1 As Long, tm2 As Long
    
    btnPingAsync.Enabled = False
    ReDim saAdresses(LV_IPs.ListItems.Count)
    For n = 1 To LV_IPs.ListItems.Count
        '//Store value
        If LV_IPs.ListItems(n).SubItems(1) = "---" Then
            LV_IPs.ListItems(n).ListSubItems(1).Tag = -1
        Else
            If Len(LV_IPs.ListItems(n).SubItems(1)) Then
                LV_IPs.ListItems(n).ListSubItems(1).Tag = LV_IPs.ListItems(n).ListSubItems(1).Text
            End If
        End If
        LV_IPs.ListItems(n).SubItems(1) = ""
        saAdresses(n - 1) = LV_IPs.ListItems(n).Text
    Next
    tm1 = GetTickCount()
    m_clsPingBase.NumParalellActions = 1000
    m_clsPingBase.PingHostList saAdresses, 100, (LV_IPs.ListItems.Count * 100) / 2
    btnPingAsync.Enabled = True
    tm2 = GetTickCount()
    m_bFirstRun = True
'    MsgBox Format((tm2 - tm1) / 1000, "0.00") & " secs", , "Time needed"
End Sub

Private Sub btnPingList_Click()
Dim n As Long, res As Long
Dim tm1 As Long, tm2 As Long
    tm1 = GetTickCount()
    btnPingList.Enabled = False
    For n = 1 To LV_IPs.ListItems.Count
        LV_IPs.ListItems(n).Selected = True
        LV_IPs.ListItems(n).EnsureVisible
        res = m_clsPingBase.PingHostSingle(LV_IPs.ListItems(n).Text, 100)
        If res < 0 Then
            LV_IPs.ListItems(n).SubItems(1) = "---"
        Else
            LV_IPs.ListItems(n).SubItems(1) = res
        End If
    Next
    btnPingList.Enabled = True
    tm2 = GetTickCount()
    MsgBox Format((tm2 - tm1) / 1000, "0.00") & " secs", , "Time needed"
End Sub

Private Sub Form_Load()
Dim n As Integer, m As Integer, k As Integer
Dim lfd As Long, sreturn As String
Dim sxIPfrom() As String, sxIPTo() As String
Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, sData As String
Dim BoundsL() As Integer, BoundsH() As Integer
    lblPResult.Caption = ""
    m_bFirstRun = False
'    For m = 1 To 15
'        For n = 1 To 255
'            LV_IPs.ListItems.Add , , "10.49." & m & "." & n
'        Next
'    Next

    lfd = FreeFile
    Open App.Path & "\iplist.txt" For Input As #lfd
        Do While Not EOF(lfd)
            Line Input #lfd, sreturn
            If Len(sreturn) Then
                If Left(sreturn, 1) <> "#" Then
                    k = InStr(sreturn, "-")
                    m = InStr(sreturn, ":")
                    If k Then   '//Komplete network
                        sxIPfrom = Split(Left(sreturn, k - 1), ".")
                        sxIPTo = Split(Mid(sreturn, k + 1), ".")
                        ReDim Preserve sxIPfrom(3)
                        ReDim Preserve sxIPTo(3)
                        ReDim BoundsH(3)
                        ReDim BoundsL(3)
                        For n = 0 To 3
                            BoundsL(n) = Val(sxIPfrom(n))
                            If BoundsL(n) > 255 Then BoundsL(n) = 255
                            BoundsH(n) = Val(sxIPTo(n))
                            If BoundsH(n) > 255 Then BoundsH(n) = 255
                        Next
                        For i1 = BoundsL(0) To BoundsH(0)
                            For i2 = BoundsL(1) To BoundsH(1)
                                For i3 = BoundsL(2) To BoundsH(2)
                                    For i4 = BoundsL(3) To BoundsH(3)
                                        sData = i1 & "." & i2 & "." & i3 & "." & i4
                                        LV_IPs.ListItems.Add , , sData
                                    Next
                                Next
                            Next
                        Next
                    ElseIf m Then
                        sxIPfrom = Split(Left(sreturn, m - 1), ".")   '//The Base adress
                        ReDim Preserve sxIPfrom(2)
                        sxIPTo = Split(Mid(sreturn, m + 1), ",")
                        For n = 0 To UBound(sxIPTo)
                            sData = sxIPfrom(0) & "." & sxIPfrom(1) & "." & sxIPfrom(2) & "." & sxIPTo(n)
                            LV_IPs.ListItems.Add , , sData
                        Next
                    Else
                        n = InStr(sreturn, " ")
                        m = InStr(sreturn, vbTab)
                        If n Then
                            sreturn = Left(sreturn, n - 1)
                        ElseIf m Then
                            sreturn = Left(sreturn, m - 1)
                        End If
                        LV_IPs.ListItems.Add , , sreturn
                    End If
                End If
            End If
        Loop
    Close #lfd
    lblHosts(1) = LV_IPs.ListItems.Count
    Set m_clsPingBase = New ClassPingBase
End Sub

Private Sub Form_Resize()
Dim x As Single
    x = Me.ScaleHeight - LV_IPs.Top - 150
    If x > 0 Then LV_IPs.Height = x
End Sub

Private Sub m_clsPingBase_PingFail(sIPAdress As String, lNewStatus As Long, ArrayIndex As Long)
    LV_IPs.ListItems(ArrayIndex + 1).SubItems(1) = "---"
    LV_IPs.ListItems.Item(ArrayIndex + 1).SmallIcon = 2
    If m_bFirstRun Then
        If LV_IPs.ListItems(ArrayIndex + 1).ListSubItems(1).Tag <> -1 Then
            '//Notify state change
            LV_IPs.ListItems(ArrayIndex + 1).ListSubItems(1).ReportIcon = 4
        End If
    End If
End Sub

'Private Sub m_clsPingBase_PingFail(sIPAdress As String, lNewStatus As Long)
'Dim lvitem As ListItem
'    Set lvitem = LV_IPs.FindItem(sIPAdress)
'    If Not lvitem Is Nothing Then
'        lvitem.SubItems(1) = "--"
'        lvitem.EnsureVisible
'    End If
'End Sub
'
'Private Sub m_clsPingBase_PingSuccess(sIPAdress As String, lNewStatus As Long)
'Dim lvitem As ListItem
'    Set lvitem = LV_IPs.FindItem(sIPAdress)
'    If Not lvitem Is Nothing Then
'        lvitem.SubItems(1) = lNewStatus
'        lvitem.EnsureVisible
'    End If
'End Sub
Private Sub m_clsPingBase_PingSuccess(sIPAdress As String, lNewStatus As Long, ArrayIndex As Long)
    LV_IPs.ListItems(ArrayIndex + 1).SubItems(1) = lNewStatus
    LV_IPs.ListItems(ArrayIndex + 1).EnsureVisible
    LV_IPs.ListItems.Item(ArrayIndex + 1).SmallIcon = 1
    If m_bFirstRun Then
        If LV_IPs.ListItems(ArrayIndex + 1).ListSubItems(1).Tag = -1 Then
            '//Notify state change
            LV_IPs.ListItems(ArrayIndex + 1).ListSubItems(1).ReportIcon = 3
        End If
    
    End If
End Sub
