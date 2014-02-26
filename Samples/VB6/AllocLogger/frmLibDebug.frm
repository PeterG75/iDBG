VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmLibDebug 
   Caption         =   "Global/Local  Alloc/Free Logger"
   ClientHeight    =   8235
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   17145
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   17145
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMsgs 
      BackColor       =   &H80000000&
      Caption         =   "Messages"
      Height          =   255
      Left            =   8640
      TabIndex        =   48
      Top             =   720
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   255
      Left            =   1920
      TabIndex        =   47
      Top             =   720
      Width           =   1455
   End
   Begin VB.CheckBox chkCrash 
      Caption         =   "Crash Check"
      Height          =   255
      Left            =   8640
      TabIndex        =   46
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   12120
      TabIndex        =   45
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame frmAdvOptions 
      Caption         =   "Advanced Options "
      Height          =   4695
      Left            =   5880
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   8295
      Begin VB.TextBox txtAdvRetInModule 
         Height          =   285
         Left            =   2400
         TabIndex        =   44
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Test"
         Height          =   375
         Left            =   6720
         TabIndex        =   42
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtAdvTest 
         Height          =   285
         Index           =   1
         Left            =   5760
         TabIndex        =   41
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtAdvTest 
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   40
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtAdvBetween 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   38
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtAdvBetween 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   37
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox chkAdvBetween 
         Caption         =   "Between"
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelFile 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   7560
         TabIndex        =   35
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtAdvInject 
         Height          =   285
         Left            =   2400
         OLEDropMode     =   1  'Manual
         TabIndex        =   34
         Top             =   3240
         Width           =   4935
      End
      Begin VB.TextBox txtAdvRetAddr 
         Height          =   285
         Left            =   2400
         TabIndex        =   32
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtAdvLessThan 
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkAdvLessThan 
         Caption         =   "Less Than"
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtAdvGreaterThan 
         Height          =   285
         Left            =   1920
         TabIndex        =   28
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox chkAdvGreaterThan 
         Caption         =   "Greater Than"
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtAdvEqualTo 
         Height          =   285
         Left            =   1920
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkAdvEqualTo 
         Caption         =   "Equal to"
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "About"
         Height          =   375
         Left            =   6615
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "ret addr in module:"
         Height          =   255
         Left            =   960
         TabIndex        =   43
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Test Config        BufSize       RetAdr"
         Height          =   375
         Left            =   3840
         TabIndex        =   39
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Inject this dll on startup/attach"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Only log if Alloc ret addr = 0x"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Only log if buffer size is (all vals in hex) "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.CheckBox chkVirtual 
      Caption         =   "VirtualAlloc/Free"
      Height          =   255
      Left            =   6840
      TabIndex        =   21
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox chkHeap 
      Caption         =   "HeapAlloc/Free"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox chkLocal 
      Caption         =   "LocalAlloc/Free"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox chkGlobal 
      Caption         =   "GlobalAlloc/Free"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "Attach"
      Height          =   375
      Left            =   10800
      TabIndex        =   17
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause/Resume"
      Height          =   375
      Left            =   12120
      TabIndex        =   15
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelFile 
      Caption         =   "..."
      Height          =   375
      Index           =   0
      Left            =   10200
      TabIndex        =   14
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Terminate"
      Height          =   375
      Left            =   13680
      TabIndex        =   13
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save to File"
      Height          =   375
      Left            =   15720
      TabIndex        =   12
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Data"
      Height          =   375
      Left            =   15720
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin Project1.ucHexEdit he 
      Height          =   5295
      Left            =   5640
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9340
   End
   Begin Project1.CmnDlg CmnDlg1 
      Left            =   0
      Top             =   405
      _ExtentX        =   582
      _ExtentY        =   503
   End
   Begin VB.TextBox txtData 
      Height          =   6975
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1080
      Width           =   11655
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7455
      Left            =   5160
      TabIndex        =   8
      Top             =   720
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13150
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Text"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Hex"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced Logging Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv 
      Height          =   7215
      Left            =   3600
      TabIndex        =   6
      Top             =   945
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   12726
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   0
      Width           =   8895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy Log"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch"
      Height          =   375
      Left            =   10800
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   7080
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "(freed)"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "hMem Objs"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Log"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Exe (DragnDrop)"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
   End
End
Attribute VB_Name = "frmLibDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: david zimmer (dzzie@yahoo.com)
'License GPL
'Date: Feb 17 2010
'
'Note: this uses the GPL iDbg debugging library (activex control) i wrote back at iDefense.com
'      idbg also uses code from oleh's ollydbg see iDbg source/credits or about button for more info.


'ideas to speed up.
'log messages to file or have different message levels to reduce logging.
'in the tracking code, we track everything, and only show it if it meets criteria
'   we should probably only log it at all if it meets the criteria (lower overhead) but could get lost.
'this was designed to deal with programs on startup and log everything, but now has attach, corner cases?
'the alloc matching cycles through each saved alloc manually in vb code..use collection key lookup if possible

'can be left in a weird state sometimes especiall when in ide

Public WithEvents dbg As CDebugger
Attribute dbg.VB_VarHelpID = -1

Dim GlobalAlloc As Long
Dim GlobalFree As Long
Dim LocalAlloc As Long
Dim LocalFree As Long
Dim HeapAlloc As Long
Dim HeapFree As Long
Dim VirtualAlloc As Long
Dim VirtualFree As Long
Dim ExitProcess As Long

Dim global_changed As Boolean
Dim local_changed As Boolean
Dim heap_changed As Boolean
Dim virtual_changed As Boolean


Dim mem As New Collection
Dim cur_mem As CMem
Dim selLi As ListItem
Dim paused As Boolean

Sub dp(msg)
    If chkMsgs.value = 0 Then Exit Sub
    List1.AddItem msg
    List1.Refresh
    Me.Refresh
    DoEvents
End Sub

'these can only be set at a breakpoint ? forget
Private Sub chkGlobal_Click()
    global_changed = True
    dbg.SuspendThreads
    initBpx GlobalAlloc, IIf(chkGlobal.value = 1, True, False)
    initBpx GlobalFree, IIf(chkGlobal.value = 1, True, False)
    dbg.ResumeThreads
End Sub

Private Sub chkHeap_Click()
    heap_changed = True
    dbg.SuspendThreads
    initBpx HeapAlloc, IIf(chkHeap.value = 1, True, False)
    initBpx HeapFree, IIf(chkHeap.value = 1, True, False)
    dbg.ResumeThreads
End Sub

Private Sub chkLocal_Click()
    local_changed = True
    dbg.SuspendThreads
    initBpx LocalAlloc, IIf(chkLocal.value = 1, True, False)
    initBpx LocalFree, IIf(chkLocal.value = 1, True, False)
    dbg.ResumeThreads
End Sub

Private Sub chkVirtual_Click()
    virtual_changed = True
    dbg.SuspendThreads
    initBpx VirtualAlloc, IIf(chkVirtual.value = 1, True, False)
    initBpx VirtualFree, IIf(chkVirtual.value = 1, True, False)
    dbg.ResumeThreads
End Sub


Private Sub cmdAttach_Click()
    
    On Error Resume Next
        
    If Not OptionsOk Then
        MsgBox "Set the logging options first", vbInformation
        Exit Sub
    End If
    
    Dim cp As CProcess
    If dbg.SelectProcess(cp) Then
        If dbg.Attach(cp.pid) Then
            dp "Attached to " & cp.pid & " successfully..."
            Text1 = "ATTACHED TO " & cp.path
        Else
            dp "Attach to " & cp.pid & " failed..."
        End If
    End If
    
End Sub

Private Sub cmdCopy_Click()
    
    On Error Resume Next
    
    Clipboard.Clear
    If txtData.Visible Then
        Clipboard.SetText txtData.Text
        MsgBox "Text Report Copied", vbInformation
    Else
        Dim c As CMem
        Set c = selLi.Tag
        Clipboard.SetText c.FreedData
        MsgBox "Binary Data Copied", vbInformation
    End If
    
End Sub

Private Sub cmdPause_Click()
    If paused Then
        dbg.ResumeThreads
        setPause False
    Else
        dbg.SuspendThreads
        setPause True
    End If
End Sub

Sub setPause(x As Boolean)
    paused = x
    cmdPause.Caption = IIf(paused, "Resume", "Pause")
End Sub

Private Sub cmdSaveAs_Click()

    On Error Resume Next
    
    Dim f As String
    f = CmnDlg1.ShowOpen("")
    If Len(f) = 0 Then Exit Sub
    
    If txtData.Visible Then
        fso.WriteFile f, txtData.Text
        MsgBox "Text Report Saved", vbInformation
    Else
        Dim c As CMem
        Set c = selLi.Tag
        fso.WriteFile f, c.FreedData
        MsgBox "Binary Data Saved", vbInformation
    End If
    
    
End Sub

Function OptionsOk() As Boolean
    If chkGlobal.value = 1 Or chkHeap.value = 1 Or chkLocal.value = 1 Or chkVirtual.value = 1 Then OptionsOk = True
End Function

Private Sub cmdSelFile_Click(Index As Integer)
   On Error Resume Next
   Dim f As String
   f = CmnDlg1.ShowOpen("", exeFiles, "Select executable file")
   If Index = 0 Then Text1 = f Else txtAdvInject = f
End Sub

Private Sub cmdLaunch_Click()
    Dim e As String
    Dim pth As String
    
    On Error Resume Next
        
    If Not OptionsOk Then
        MsgBox "Set the logging options first", vbInformation
        Exit Sub
    End If
        
        
    cmdLaunch.Enabled = False

    List1.Clear
    lv.ListItems.Clear
    Set mem = New Collection
    Set selLi = Nothing
    Set cur_mem = Nothing
    txtData = Empty
        
    pth = Text1
    If Dir(pth) = "" Then
        MsgBox "Path not found " & pth
        Exit Sub
    End If
    
    setPause False
    If Not dbg.LaunchProcess(pth) Then
        MsgBox "launch error:" & dbg.GetErr
        cmdLaunch.Enabled = True
    Else
        dp "starting"
    End If
    
End Sub



Private Sub Command1_Click()
    MsgBox "Alloc/Free Logger - Feb 17 - 2010" & vbCrLf & _
            "Author: David Zimmer (dzzie@yahoo.com)" & vbCrLf & _
            "License GPL" & vbCrLf & "Uses iDefense iDbg Debugger Library - Credits Follow", vbInformation
    dbg.About
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim i As Long
    Dim tmp As String
    
    For i = 0 To List1.ListCount
        tmp = tmp & List1.List(i) & vbCrLf
    Next
    
    Clipboard.Clear
    Clipboard.SetText tmp
        
End Sub

Private Function ForcedExtraction(Optional lli As ListItem = Nothing)
    Dim li As ListItem
    Dim c As CMem
    Dim buf() As Byte
    
    If lli Is Nothing Then
        For Each c In mem
            If c.FreedFrom = 0 Then
                ReDim buf(c.bufsize)
                If dbg.ReadBuf(c.BufAddr, c.bufsize, buf) Then
                    c.FreedData = StrConv(buf, vbUnicode)
                    c.SetLiColor vbRed
                Else
                    'dp "ForcedExtraction failed to readbuf on " & Hex(c.BufAddr)
                End If
            End If
        Next
    Else
        Set c = lli.Tag
        If c.FreedFrom = 0 Then
            ReDim buf(c.bufsize)
            If dbg.ReadBuf(c.BufAddr, c.bufsize, buf) Then
                c.FreedData = StrConv(buf, vbUnicode)
            Else
                'dp "ForcedExtraction failed to readbuf on " & Hex(c.BufAddr)
            End If
        End If
    End If
            
    
End Function
Private Sub Command3_Click()
    dbg.StopDbg
    setPause False
    cmdLaunch.Enabled = True
End Sub

Private Sub Command4_Click()
    
    Dim a As Long, b As Long, c As Boolean, e As String
    
    If Not getHex(txtAdvTest(0), a) Then
        MsgBox "Invalid hex number for test field bufsize", vbInformation
        Exit Sub
    End If
    
    If Not getHex(txtAdvTest(1), b) Then
        MsgBox "Invalid hex number for test field retaddr", vbInformation
        Exit Sub
    End If
    
    c = CheckIfVisible(a, b, e)
    
    MsgBox "CheckifVisible returned: " & c & vbCrLf & vbCrLf & "ErrMsg: " & vbCrLf & e, vbInformation
    
    
End Sub

Private Sub Command5_Click()
    List1.Clear
End Sub

Private Sub dbg_DebugString(msg As String)
    dp "Debug string: " & msg
End Sub

Private Sub dbg_DevMessage(func As String, msg As String)
    
        'List1.AddItem "DevMsg: " & func & " " & msg

End Sub

Private Sub dbg_DllLoad(path As String, base As Long, size As Long)
    'dp "Loading dll " & path & " at base " & Hex(base)
End Sub

Private Sub dbg_Exception(except As CException) 'you must call continue/step* to resume
  
        
       dp "Exception code " & Hex(except.ExceptionCode) & " at addr 0x" & Hex(except.ExceptionAddress)
'
        If except.ExceptionCode = dbg.LastEventCode And _
            except.ExceptionAddress = dbg.LastEventAddress And chkCrash.value = 1 Then
                dp "Same crash killing"
                ForcedExtraction
                dbg.StopDbg
        Else
                dbg.Continue DBG_EXCEPTION_NOT_HANDLED
        End If
        
   
   
End Sub


Private Function resolve(fxName As String, fxVar As Long, Optional libName As String = "kernel32") As Boolean
    fxVar = dbg.ResolveExport(libName, fxName)
    dp "Resolved: " & fxName & " = 0x" & Hex(fxVar)
End Function


Private Function initBpx(fxVar As Long, Optional setit As Boolean = True) As Boolean

    On Error GoTo hell
    
    If setit Then
       If Not dbg.SetBreakPoint(fxVar) Then
           'dp "Error adding " & fxName & " breakpoint"
       Else
           'dp "Set " & fxName & " bpx @ " & Hex(fxVar)
           initBpx = True
       End If
    Else
        If Not dbg.RemoveBreakpoint(fxVar) Then
          ' dp "Error removing " & fxName & " breakpoint"
       'Else
       '    dp "Removed " & fxName & " bpx @ " & Hex(fxVar)
       '    initBpx = True
       End If
   End If
   
hell:
   
End Function


Private Sub dbg_ReadyToRun()

    Dim e As String
    Dim bp As Long
    
    dp "ReadyToRun...Adding startup bpxs"
    
    resolve "GlobalAlloc", GlobalAlloc
    resolve "GlobalFree", GlobalFree
    resolve "LocalAlloc", LocalAlloc
    resolve "LocalFree", LocalFree
    resolve "RtlAllocateHeap", HeapAlloc, "ntdll"
    resolve "RtlFreeHeap", HeapFree, "ntdll"
    resolve "VirtualAlloc", VirtualAlloc
    resolve "VirtualFree", VirtualFree
    resolve "ExitProcess", ExitProcess
    
    initBpx GlobalAlloc, IIf(chkGlobal.value = 1, True, False)
    initBpx GlobalFree, IIf(chkGlobal.value = 1, True, False)
    initBpx LocalAlloc, IIf(chkLocal.value = 1, True, False)
    initBpx LocalFree, IIf(chkLocal.value = 1, True, False)
    initBpx HeapAlloc, IIf(chkHeap.value = 1, True, False)
    initBpx HeapFree, IIf(chkHeap.value = 1, True, False)
    initBpx VirtualAlloc, IIf(chkVirtual.value = 1, True, False)
    initBpx VirtualFree, IIf(chkVirtual.value = 1, True, False)
    
    initBpx ExitProcess
              
    If fso.FileExists(txtAdvInject) Then
        dp "Injecting: " & fso.FileNameFromPath(txtAdvInject) & " ok? " & dbg.InjectDLL(txtAdvInject)
    End If
    
    dbg.DbgContinue
    
End Sub

Private Sub dbg_SingleStep(except As CException)
     dp "Single step at addr " & Hex(except.ExceptionAddress) & " disasm = " & except.Disasm & " return value from call was (eax) = " & Hex(dbg.ReadRegister(Eax))
     dbg.Continue
End Sub

Private Sub dbg_Terminate()
    dp "Terminated"
    setPause False
    cmdLaunch.Enabled = True
End Sub

Private Sub dbg_ThreadCreate(threadID As Long, startAddress As Long)
    'dp "Thread Created: " & threadID & " at " & Hex(startAddress)
End Sub

Private Sub dbg_ThreadDestroy(threadID As Long, exitCode As Long)
    'dp "Thread " & threadID & " exited with code " & exitCode
End Sub

Private Sub dbg_UserBreakpoint(except As CException) 'you must call continue/step* to resume
    Dim l As Long
    Dim retadr As Long
    Dim arg8 As Long
    Dim arg4 As Long
    Dim argC As Long
    Dim validPointer As Boolean
    Dim amod As String
    Dim aproc As String
    Dim buf() As Byte
    Dim li As ListItem
    Dim c As CMem
    
    'at function start esp=retaddr, esp+4 is arg1, esp+8 is arg2, esp+C = arg3
    'Global/LocalAlloc(Flags, size)
    'GLobal/LocalFree(hMem)
    'HeapAlloc(hHeap,dwFlags,dwBytes); //just forwards to ntdll.RtlAllocateHeap
    'HeapFree(hHeap,dwFlags,lpMem);    //just forwards to ntdll.RtlFreeHeap
    'HeapReAlloc(hHeap,dwFlags,lpMem,dwBytes);
    'VirtualFree(lpAddress,dwSize,dwFreeType);
    'VirtualAlloc(lpAddress,dwSize,flAllocationType,flProtect);
    
    'this block is probably not of functional use right now
'    If global_changed Then
'        initBpx GlobalAlloc, IIf(chkGlobal.value = 1, True, False)
'        initBpx GlobalFree, IIf(chkGlobal.value = 1, True, False)
'        global_changed = False
'    End If
'
'    If heap_changed Then
'        initBpx HeapAlloc, IIf(chkHeap.value = 1, True, False)
'        initBpx HeapFree, IIf(chkHeap.value = 1, True, False)
'        heap_changed = False
'    End If
'
'    If local_changed Then
'        initBpx LocalAlloc, IIf(chkLocal.value = 1, True, False)
'        initBpx LocalFree, IIf(chkLocal.value = 1, True, False)
'        local_changed = False
'    End If
'
'    If virtual_changed Then
'        initBpx VirtualAlloc, IIf(chkVirtual.value = 1, True, False)
'        initBpx VirtualFree, IIf(chkVirtual.value = 1, True, False)
'        virtual_changed = False
'    End If
    

'cur_mem is global variable becauses alloc sets bpx on ret to log result.
'maybe a hook implementation is better...

    If except.ExceptionAddress = GlobalAlloc Or _
        except.ExceptionAddress = LocalAlloc Or _
         except.ExceptionAddress = VirtualAlloc Then
         
        'these all give the size as second argument which is what we are logging.
        'VirtualAlloc(lpAddress,dwSize,flAllocationType,flProtect);
        'Global/LocalAlloc(Flags, size)
        
        Set cur_mem = New CMem
        
        Select Case except.ExceptionAddress
            Case GlobalAlloc:  cur_mem.FunctionName = "GlobalAlloc"
            Case LocalAlloc:   cur_mem.FunctionName = "LocalAlloc"
            Case VirtualAlloc: cur_mem.FunctionName = "VirtualAlloc"
        End Select
        
        l = dbg.ReadRegister(esp)
        dbg.ReadLng l, retadr
        cur_mem.retaddr = retadr
        
        dbg.ReadLng l + 8, arg8
        cur_mem.bufsize = arg8
        
        dp cur_mem.FunctionName & " ret addr=" & Hex(retadr) & " sz=0x" & Hex(arg8) & " (" & arg8 & "d)"
        dbg.SetBreakPoint retadr, True 'set as oneshot (love that feature!)
        
    ElseIf except.ExceptionAddress = GlobalFree Or _
            except.ExceptionAddress = LocalFree Or _
             except.ExceptionAddress = VirtualFree Then
             
        'these all give the free'd address as the first argument
        l = dbg.ReadRegister(esp)
        dbg.ReadLng l, retadr
        dbg.ReadLng l + 4, arg4
        
        If Not HandleDataFree(arg4, retadr) Then
            NotFound arg4, retadr
            'dp "Free Failed to find hMem 0x" & Hex(arg4)
        End If
        
    ElseIf except.ExceptionAddress = HeapAlloc Then
        'HeapAlloc(hHeap,dwFlags,dwBytes);
        Set cur_mem = New CMem
        cur_mem.FunctionName = "HeapAlloc"
        
        l = dbg.ReadRegister(esp)
        dbg.ReadLng l, retadr
        cur_mem.retaddr = retadr
        
        dbg.ReadLng l + &HC, argC
        cur_mem.bufsize = argC
        
        dp cur_mem.FunctionName & " ret addr=" & Hex(retadr) & " sz=0x" & Hex(argC) & " (" & argC & "d)"
        dbg.SetBreakPoint retadr, True 'set as oneshot (love that feature!)
        
    ElseIf except.ExceptionAddress = HeapFree Then
        
        l = dbg.ReadRegister(esp) 'HeapFree(hHeap,dwFlags,lpMem);
        dbg.ReadLng l, retadr
        dbg.ReadLng l + &HC, argC
        
        If Not HandleDataFree(argC, retadr) Then
            NotFound argC, retadr
            'dp "Free Failed to find hMem 0x" & Hex(argC)
        End If
    
    Else
        
        'its the address xxAlloc was supposed to return to
        If Not (cur_mem Is Nothing) And except.ExceptionAddress = cur_mem.retaddr Then
            
            'If cur_mem.bufsize <> 0 And cur_mem.BufAddr <> 0 Then
            
                cur_mem.BufAddr = dbg.ReadRegister(Eax)
                mem.Add cur_mem
        
                cur_mem.Visible = CheckIfVisible(cur_mem.bufsize, cur_mem.retaddr)
                
                If cur_mem.Visible Then
                    Set li = lv.ListItems.Add(, , "0x" & Hex(cur_mem.BufAddr) & " (0x" & Hex(cur_mem.bufsize) & ")")
                    Set li.Tag = cur_mem
                    Set cur_mem.li = li
                End If
            
            'End If
            
            Set cur_mem = New CMem
            
        ElseIf except.ExceptionAddress = ExitProcess Then
            ForcedExtraction
        Else
            dp "Unexpected bpx hit: @ 0x" & Hex(except.ExceptionAddress)
        End If

    End If
    
end_of_func:
            
    dbg.Continue
    
End Sub

Private Function HandleDataFree(freeAddr As Long, retadr As Long) As Boolean
    Dim c As CMem
    Dim li As ListItem
    Dim buf() As Byte
    
    'this is a slow way to do this probably, but safer than mem collection key lookup (dups)
       For Each c In mem
            If c.BufAddr = freeAddr Then
                ReDim buf(c.bufsize)
                c.FreedFrom = retadr
                If dbg.ReadBuf(freeAddr, c.bufsize, buf()) Then
                    c.FreedData = StrConv(buf, vbUnicode)
                    dp c.FunctionName & "(" & Hex(c.BufAddr) & ")=" & c.FreedData
                    c.SetLiColor vbBlue
                    HandleDataFree = True
                Else
                    dp "Failed to readbuf for " & c.FunctionName & "(" & Hex(c.BufAddr) & ")"
                End If
            End If
        Next
        
End Function

'this is used if we attach to a process and didnt catch the alloc
Private Function NotFound(freeAddr As Long, retadr As Long)
    Dim c As CMem
    Dim buf() As Byte
    ReDim buf(&H51)
    
    If freeAddr = 0 Then Exit Function
    
    Set c = New CMem
    c.NotFound = True
    c.bufsize = &H51
    c.BufAddr = freeAddr
    c.FreedFrom = retadr
    c.FunctionName = "Alloc Not Found (Attach?)"
    
    If dbg.ReadBuf(freeAddr, c.bufsize, buf()) Then
        c.FreedData = StrConv(buf, vbUnicode)
        dp c.FunctionName & "(" & Hex(c.BufAddr) & ")=" & c.FreedData
    Else
        dp "Failed to readbuf for " & c.FunctionName & "(" & Hex(c.BufAddr) & ")"
    End If
    
    mem.Add c
    c.Visible = True ' CheckIfVisible(c.bufsize, c.retaddr)
            
    If c.Visible Then
        Set li = lv.ListItems.Add(, , "0x" & Hex(cur_mem.BufAddr) & " (0x" & Hex(cur_mem.bufsize) & ")")
        Set li.Tag = c
        Set c.li = li
        c.SetLiColor vbCyan
    End If
            
    
        
End Function


Private Sub Form_Load()
    mnuPopup.Visible = False
    he.Move txtData.Left, txtData.Top, txtData.Width, txtData.Height
    frmAdvOptions.Move txtData.Left, txtData.Top, txtData.Width, txtData.Height
    Text1 = App.path & "\test.exe"
    Set dbg = New CDebugger
    dbg.UseSymbols = True
    dbg.SymbolPath = "c:\"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  If dbg.isDebugging Then dbg.StopDbg
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim c As CMem
    
    If Item.ForeColor <> vbBlue Then
        ForcedExtraction Item
    End If
    
    Set selLi = Item
    Set c = Item.Tag
    
    txtData = c.GetReport
       
    Dim f As Long
    Const tmp = "c:\tmp.tmp"
    If fso.FileExists(tmp) Then Kill tmp
    f = FreeFile
    Open tmp For Binary As f
    Put f, , c.GetDataBytes
    Close f
    
    he.LoadFile tmp
    Kill tmp
    
End Sub

Private Function CheckIfVisible(bufsize As Long, retaddr As Long, Optional errmsg As String) As Boolean
    
    Dim ret As Long
    Dim an_opt_is_set As Boolean
    Dim should_log As Boolean
    Dim ret2 As Long
    Dim e() As String
    Dim m As String
    
    an_opt_is_set = False
    should_log = False
    
    If Len(txtAdvRetAddr) > 0 Then
        If getHex(txtAdvRetAddr, ret) Then
                an_opt_is_set = True
                If ret = retaddr Then
                    should_log = True
                    push e, "ReturnAddress Match Made"
                End If
        Else
            push e, "RetAddr is invalid hex number: " & txtAdvRetAddr
        End If
    End If
    
    If chkAdvGreaterThan.value = 1 And Len(txtAdvGreaterThan) > 0 Then
        If getHex(txtAdvGreaterThan, ret) Then
                an_opt_is_set = True
                If bufsize >= ret Then
                    should_log = True
                    push e, "Greater than Match Made"
                End If
        Else
            push e, "Greater Than is invalid hex number: " & txtAdvGreaterThan
        End If
    End If
            
    If chkAdvLessThan.value = 1 And Len(txtAdvLessThan) > 0 Then
        If getHex(txtAdvLessThan, ret) Then
            'If ret > 0 Then
                an_opt_is_set = True
                If bufsize <= ret Then
                    should_log = True
                    push e, "less than Match Made"
                End If
            'End If
         Else
            push e, "Less Than is invalid hex number: " & txtAdvLessThan
        End If
    End If
    
    If chkAdvEqualTo.value = 1 And Len(txtAdvEqualTo) > 0 Then
        If getHex(txtAdvEqualTo, ret) Then
            'If ret > 0 Then
                an_opt_is_set = True
                If bufsize = ret Then
                    should_log = True
                    push e, "Equal to Match Made"
                End If
            'End If
         Else
            push e, "Equal tois invalid hex number: " & txtAdvEqualTo
        End If
    End If
    
    If chkAdvBetween.value = 1 And Len(txtAdvBetween(0)) > 0 And Len(txtAdvBetween(1)) > 0 Then
        If getHex(txtAdvBetween(0), ret) Then
            'If ret > 0 Then
                If getHex(txtAdvBetween(1), ret2) Then
                    'If ret2 > 0 Then
                        an_opt_is_set = True
                        If bufsize >= ret And bufsize <= ret2 Then
                            should_log = True
                            push e, "Between Match Made"
                        End If
                    'End If
                 Else
                    push e, "Between High num is invalid hex number: " & txtAdvBetween(1)
                End If
            'End If
         Else
            push e, "Between Low Num is invalid hex number: " & txtAdvBetween(0)
        End If
    End If
    
    If Len(txtAdvRetInModule) > 0 Then
        an_opt_is_set = True
        m = dbg.ModuleAtVA(retaddr)
        If Len(m) > 0 Then
            If InStr(m, txtAdvRetInModule) > 0 Then
                should_log = True
                push e, "Ret in module match made: " & m
            End If
        Else
            push e, "Must be running for dbg.ModuleAtVa to work"
        End If
    End If
    
    errmsg = Join(e, vbCrLf)
    
    If Not an_opt_is_set Then
        CheckIfVisible = True
        Exit Function
    End If
    
    CheckIfVisible = should_log
            
End Function

Private Function getHex(s As String, ret As Long) As Boolean
    On Error GoTo hell
    ret = CLng("&h" & s)
    getHex = True
    Exit Function
hell:
    ret = 0
End Function

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuSaveAs_Click()
    If selLi Is Nothing Then Exit Sub
    
    Dim c As CMem
    Set c = selLi.Tag
    
    Dim fpath As String
    fpath = CmnDlg1.ShowSave(App.path, AllFiles, "Save As")
    If Len(fpath) = 0 Then Exit Sub
    
    Dim f As Long
     
    If fso.FileExists(fpath) Then Kill fpath
    
    f = FreeFile
    Open fpath For Binary As f
    Put f, , c.GetDataBytes
    Close f
    
End Sub

Private Sub TabStrip1_Click()
    With TabStrip1
        txtData.Visible = IIf(.SelectedItem.Index = 1, True, False)
        he.Visible = IIf(.SelectedItem.Index = 2, True, False)
        frmAdvOptions.Visible = IIf(.SelectedItem.Index = 3, True, False)
    End With
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Text1 = Data.Files(1)
End Sub

Private Sub txtAdvInject_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If fso.FileExists(Data.Files(1)) Then txtAdvInject = Data.Files(1)
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub
