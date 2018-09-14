Attribute VB_Name = "Module1"
Option Explicit

'Scans for VB6 IDE mouse over variable tooltip windows,
'extracts the text. If its an numeric value then displays its hex value in a new popup window
'right next to the original tool tip.

'Author: DllHell
'Thread: http://www.vbforums.com/showthread.php?862681-RESOLVED-IDE-mouse-over-variables-tooltip&p=5290621#post5290621

Public Const LB_SETTABSTOPS = &H192
Public Const MAX_PATH = 260

Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameA" (ByVal hProcess As Long, ByVal lpImageFileName As String, ByVal nSize As Long) As Long

Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

Const WM_SETTEXT As Long = &HC
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public found As String
Public lastRect As RECT
Public tt As New frmToolTip

Public Function ShowExtraToolTip()
    Dim hWnd As Long
    Dim f() As String
    Dim i As Long
    Dim l As Long
    Dim r As RECT
    Dim txt As String
    Dim newTxt As String
    Dim foundOne As Boolean
    Dim before As String
    Dim after As String
    
    found = ""

    'load the "found" var with a csv list of tooltip handles and their captions in the format hwnd=text,hwnd=text
    Call EnumWindows(AddressOf EnumWindowsCallBack, hWnd)
    
    'look through the results (note there may be some false positives)
    If found <> "" Then
        f = Split(found, "~~~")
        For i = LBound(f) To UBound(f)
            If f(i) <> "" Then
                hWnd = Split(f(i), "===", 2)(0)
                txt = Split(f(i), "===", 2)(1)
                
                'tool tip we are interested in have the format "var = value" so break into before and after the equal sign
                If InStr(1, txt, "=") > 0 Then
                    before = Split(txt, "=", 2)(0)
                    after = Split(txt, "=", 2)(1)
                Else
                    before = txt
                    after = ""
                End If
                
                If after <> "" Then
                    'check the 'after' var to see if you need to display your new tooltip, in this case looking for numbers
                    If IsNumeric(after) Then
                        foundOne = True
                        newTxt = " [0x" & Right("00000000" & Hex(after), 8) & "] "
                        l = GetWindowRect(hWnd, r)
                        If l <> 0 Then
                            Call tt.ShowMe(hWnd, r, newTxt)
                            foundOne = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    If Not foundOne Then
        tt.Hide
        lastRect = newRect
    End If
    
End Function

Public Function EnumWindowsCallBack(ByVal hWnd As Long, ByVal lpData As Long) As Long
    Dim lResult    As Long
    Dim lThreadId  As Long
    Dim lProcessId As Long
    Dim sWndName   As String
    Dim sClassName As String
    
    Dim tl As Long
    Dim tooltiptext As String
    
    EnumWindowsCallBack = 1 'keep enumerating
    sClassName = Space$(MAX_PATH)
    sWndName = Space$(MAX_PATH)
    
    lResult = GetClassName(hWnd, sClassName, MAX_PATH)
    sClassName = Left$(sClassName, lResult)
    lResult = GetWindowText(hWnd, sWndName, MAX_PATH)
    sWndName = Left$(sWndName, lResult)
    
    lThreadId = GetWindowThreadProcessId(hWnd, lProcessId)
    
    If UCase(sClassName) = UCase("tooltips_class32") Then 'filter on tooltips
        If ExeNameFromProcID(lProcessId) = "VB6.EXE" Then 'only tt from vb6.exe
            tl = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, 0) 'then tt with readable text
            If tl > 0 Then
                tooltiptext = Space$(tl)
                Call SendMessage(hWnd, WM_GETTEXT, ByVal tl + 1, ByVal tooltiptext)
                found = found & hWnd & "===" & tooltiptext & "~~~"
            End If
        End If
    End If
    
End Function

Private Function ExeNameFromProcID(procID As Long) As String

    Dim s As String, a As Long
    
    s = ExePathFromProcID(procID)
    s = Mid(s, InStrRev(s, "\") + 1)
    a = InStr(s, Chr(0))
    
    If a > 0 Then
        s = Mid(s, 1, a - 1)
    End If
    
    ExeNameFromProcID = UCase(s)
    
End Function

Private Function ExePathFromProcID(idProc As Long) As String
    Const MAX_PATH = 260
    Const PROCESS_QUERY_INFORMATION = &H400
    Const PROCESS_VM_READ = &H10

    Dim sBuf As String
    Dim sChar As Long, l As Long, hProcess As Long
    sBuf = String$(MAX_PATH, Chr$(0))
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, idProc)
    If hProcess Then
        sChar = GetProcessImageFileName(hProcess, sBuf, MAX_PATH)
        If sChar Then
            sBuf = Left$(sBuf, sChar)
            ExePathFromProcID = sBuf
            
        End If
        CloseHandle hProcess
    End If
End Function

Function ntrim(ByVal theString As String) As String
  Dim iPos As Long
  iPos = InStr(theString, Chr$(0))
  If iPos > 0 Then theString = Left$(theString, iPos - 1)
End Function

Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function

Private Function newRect() As RECT
'intentionally left blank, this will return a new rect struct
End Function

