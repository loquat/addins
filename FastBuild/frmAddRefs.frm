VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddRefs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fast Build - Add References"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvFiltered 
      Height          =   1950
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   3440
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
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
   Begin VB.TextBox txtDetails 
      Height          =   1050
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3675
      Width           =   5145
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   75
      TabIndex        =   0
      Top             =   3000
      Width           =   3795
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   825
         TabIndex        =   1
         Top             =   90
         Width           =   2625
      End
      Begin VB.Label Label2 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3600
         TabIndex        =   4
         Top             =   135
         Width           =   195
      End
      Begin VB.Label Label1 
         Caption         =   "Search"
         Height          =   285
         Left            =   135
         TabIndex        =   2
         Top             =   135
         Width           =   645
      End
   End
   Begin MSComctlLib.ListView lv2 
      Height          =   2850
      Left            =   75
      TabIndex        =   6
      Top             =   75
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   5027
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
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
End
Attribute VB_Name = "frmAddRefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Author: David Zimmer
' Site:   http://sandsprite.com
'
Option Explicit

Dim reg As New CReg
Dim tlbs As Collection
Dim selEntry As CEntry

Private Sub Form_Load()
   
    'With lv
    '    lv2.Move .Left, .Top, .Width, .Height
    '    lvFiltered.Move .Left + SSTab1.Left, .Top + SSTab1.Top, .Width, .Height
    'End With
    
    With lv2
        lvFiltered.Move .Left, .Top, .Width, .Height
    End With
    
    'lv.ColumnHeaders(1).Width = lv.Width
    lv2.ColumnHeaders(1).Width = lv2.Width
    lvFiltered.ColumnHeaders(1).Width = lv2.Width
    
    
    Set tlbs = New Collection
    
    'Me.Visible = True
    'Me.Refresh
    'DoEvents
    
    reg.hive = HKEY_CLASSES_ROOT
    'BuildComponentList 'you can not add component refs with the addin api apparently...
                        '(this was originally developed for the logic and LazActiveX UI)
    
    'SSTab1.Tab = 1
    'SSTab1.Enabled = False
    
    BuildReferenceList
End Sub


Function BuildReferenceList()
    
    Dim clsids() As String
    Dim clsid
    Dim li As ListItem
    Dim e As CEntry
    Dim tmp As CEntry
    Dim vers() As String
    Dim revs() As String
    Dim c As New Collection
    Dim lia As ListItem
    
    If reg.hive = HKEY_CLASSES_ROOT Then
        clsids = reg.EnumKeys("\TypeLib")
    Else
        'Stop
        'clsids = reg.EnumKeys("\SOFTWARE\Classes\CLSID")
    End If
    
    For Each clsid In clsids

        Set e = New CEntry
        e.clsid = clsid
        
        'If clsID = "{00025E01-0000-0000-C000-000000000046}" Then Stop
                
         vers() = reg.EnumKeys("\TypeLib\" & clsid)
         If AryIsEmpty(vers) Then GoTo nextone
                
         revs() = reg.EnumKeys("\TypeLib\" & clsid & "\" & vers(UBound(vers)))
         If AryIsEmpty(revs) Then GoTo nextone
         
         With e
            e.name = reg.ReadValue("\TypeLib\" & clsid & "\" & vers(UBound(vers)), "")
            e.path = reg.ReadValue("\TypeLib\" & clsid & "\" & vers(UBound(vers)) & "\" & revs(0) & "\win32", "")
            .version = vers(UBound(vers)) & "." & revs(0)
            
            e.path = ValidatePath(e.path)
            
            If FileExists(e.path) And Len(.name) > 0 Then
                If Not KeyExistsInCollection(.name, c) Then
                    Set li = lv2.ListItems.Add(, , .name)
                    Set li.Tag = e
                    If RefAlreadyExists(.clsid) Then
                        .AlreadyReferenced = True
                        li.Checked = True
                    End If
                    c.Add e, .name
                End If
            End If
            
        End With
        
nextone:

   Next
End Function

Function ValidatePath(fpath As String) As String
    
    Dim a As Long
    Dim b As Long
    'example input: C:\WINDOWS\system32\catsrvut.dll\2
    
    a = InStrRev(fpath, ".")
    b = InStrRev(fpath, "\")
    If b > a Then
        ValidatePath = Mid(fpath, 1, b - 1)
    Else
        ValidatePath = fpath
    End If
    
End Function



Function GetExtension(path) As String
    Dim tmp, ub
    If Len(path) = 0 Then Exit Function
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetExtension = LCase(Mid(ub, InStrRev(ub, "."), Len(ub)))
    Else
       GetExtension = ""
    End If
End Function

'no addin api to add ocx controls :-\
'Function BuildComponentList()
'
'    Dim clsids() As String
'    Dim clsid
'    Dim li As ListItem
'    Dim e As CEntry
'    Dim tmp As CEntry
'
'    'Const full = "\SOFTWARE\Classes\CLSID\{0DE63042-EB7E-4449-BF13-7FF73866F20E}\Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}"
'    Const catid_control = "\Implemented Categories\{40FC6ED4-2438-11cf-A3DB-080036F12502}"
'    Const catid_programmable = "\Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}"
'    Const server = "\InprocServer32"
'
'    If reg.hive = HKEY_CLASSES_ROOT Then
'        clsids = reg.EnumKeys("\CLSID")
'    Else
'        clsids = reg.EnumKeys("\SOFTWARE\Classes\CLSID")
'    End If
'
'    For Each clsid In clsids
'
'        Set e = New CEntry
'        e.clsid = clsid
'
'        'If clsID = "{66CBC149-A49F-48F9-B17A-6A3EA9B42A87}" Then Stop
'
'        If reg.hive = HKEY_CLASSES_ROOT Then
'            clsid = "\CLSID\" & clsid
'        Else
'            clsid = "\SOFTWARE\Classes\CLSID\" & clsid
'        End If
'
'        With e
'
'            .isControl = reg.keyExists(clsid & "\Control")
'            If .isControl = False Then
'                If reg.keyExists(clsid & catid_control) Then .isControl = True
'            End If
'
'            '.isProgrammable = reg.keyExists(clsID & "\Programmable")
'            'If .isProgrammable = False Then
'            '    If reg.keyExists(clsID & catid_programmable) Then .isProgrammable = True
'            'End If
'
'            .typeLib = reg.ReadValue(clsid & "\typeLib", "")
'
'            If Len(.typeLib) > 0 Then
'
'                'If e.isControl And KeyExistsInCollection(.typeLib, tlbs) Then
'                '    Set tmp = tlbs(.typeLib)
'                '    If Not tmp.isControl Then tlbs.Remove tmp.typeLib   'we will update it below..
'                'End If
'
'                'If (.isControl Or .isProgrammable) And Not KeyExistsInCollection(.typeLib, tlbs) Then
'                If .isControl And Not KeyExistsInCollection(.typeLib, tlbs) Then
'                    .name = GetName(.typeLib)
'                    If Len(.name) > 0 Then
'                        .path = reg.ReadValue(clsid & "\InprocServer32", "")
'                        .progID = reg.ReadValue(clsid & "\ProgID", "")
'                        .version = reg.ReadValue(clsid & "\version", "")
'                        tlbs.Add e, .typeLib
'
'                        If e.isControl Then
'                            Set li = lv.ListItems.Add(, , .name)
'                            Set li.Tag = e
'                            .AlreadyReferenced = RefAlreadyExists(.clsid)
'                            li.Checked = .AlreadyReferenced
'                        'ElseIf e.isProgrammable Then
'                        '    Set li = lv2.ListItems.Add(, , .name)
'                        '    Set li.Tag = e
'                        End If
'                    End If
'
'                End If
'            End If
'
'        End With
'
'   Next
'End Function

Function ExistsInLV(s, lv As ListView) As Boolean
    Dim li As ListItem
    For Each li In lv.ListItems
        If li.text = s Then ExistsInLV = True
    Next
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function KeyExistsInCollection(key As String, c As Collection) As Boolean
    On Error GoTo hell
    Dim x
    Set x = c(key)
    KeyExistsInCollection = True
    Exit Function
hell:
End Function

Function GetName(typeLibID As String)
    Dim keys() As String
    Dim k, v As String, base As String

    If reg.hive = HKEY_CLASSES_ROOT Then
        base = "\TypeLib\" & typeLibID
    Else
        base = "\SOFTWARE\Classes\TypeLib\" & typeLibID
    End If
    
    keys() = reg.EnumKeys(base)
    If AryIsEmpty(keys) Then Exit Function
    For Each k In keys
       v = reg.ReadValue(base & "\" & k, "")
       If Len(v) > 0 Then
           GetName = v
           Exit Function
       End If
    Next
End Function



Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
 

Private Sub Label2_Click()
    MsgBox "Case insensitive search. Type 'checked' to see active references", vbInformation
End Sub

Private Sub lvFiltered_ItemCheck(ByVal item As MSComctlLib.ListItem)
    
    Dim li As ListItem
    Dim llv As ListView
    
    'we need to sync parent list..
    'If SSTab1.Tab = 0 Then Set llv = lv Else Set llv = lv2
    Set llv = lv2
    
    For Each li In llv.ListItems
        If li.text = item.text Then
            li.Checked = True 'apparently this doesnt fire the _ItemCheck event like it thought..
            Exit For
        End If
    Next
    
    HandleItemCheck item
    
End Sub

Private Sub lv2_ItemCheck(ByVal item As MSComctlLib.ListItem)
    HandleItemCheck item
End Sub

'Private Sub lv_ItemCheck(ByVal item As MSComctlLib.ListItem)
'    HandleItemCheck item
'End Sub

Sub HandleItemCheck(ByVal item As MSComctlLib.ListItem)

    On Error GoTo hell
    
    Set selEntry = item.Tag
    txtDetails = selEntry.ToString()
    
    If selEntry Is Nothing Then Exit Sub
       
    Dim r As Reference
    Dim guid As String
    
    guid = selEntry.clsid
    
    If item.Checked Then
        'If SSTab1.Tab = 0 Then 'components
        '    VBInstance.ActiveVBProject.AddToolboxProgID selEntry.progID
        '    selEntry.AlreadyReferenced = True
        'Else
            Set r = VBInstance.ActiveVBProject.References.AddFromFile(selEntry.path)
            selEntry.AlreadyReferenced = True
            If r Is Nothing Then
                MsgBox "Could not add reference to " & guid
                Exit Sub
            End If
        'End If
    Else
        
        'If SSTab1.Tab = 0 Then
        '    MsgBox "Sorry i cant remove components from the toolbox?", vbInformation
        '    item.Checked = True
        '    Exit Sub
        'End If
        
        'if you remove ref for an ocx
        'it wont remove from toolbox..if then use ide to remove compoenent crash here:
        '004A21CB  cmp         word ptr [ecx+32h],0
        
        'If SSTab1.Tab = 0 Then guid = selEntry.progID
        Set r = GetReference(guid) ',  (SSTab1.Tab = 0))
        If r Is Nothing Then
            MsgBox "Could not find reference to " & guid
            Exit Sub
        End If
        'this can fail for default references..a boolean return would have been nice..
        'we should recheck getreference and recheck box if it failed..but to lazy for small bug..
       
        VBInstance.ActiveVBProject.References.Remove r
        selEntry.AlreadyReferenced = False
         
        
    End If
    
    
    Exit Sub
hell:
    MsgBox "Error: " & Err.Description
    
End Sub

'Private Sub lv_ItemClick(ByVal item As MSComctlLib.ListItem)
'    Set selEntry = item.Tag
'    txtDetails = selEntry.ToString()
'End Sub

Private Sub lv2_ItemClick(ByVal item As MSComctlLib.ListItem)
    Set selEntry = item.Tag
    txtDetails = selEntry.ToString()
End Sub

Private Sub lvFiltered_ItemClick(ByVal item As MSComctlLib.ListItem)
    Set selEntry = item.Tag
    txtDetails = selEntry.ToString()
End Sub


'Private Sub SSTab1_Click(PreviousTab As Integer)
'
'    'load on demand to reduce startup time
'    If SSTab1.Tab = 1 And lv2.ListItems.Count = 0 Then BuildReferenceList
'
'    If Len(txtSearch) > 0 Then txtSearch_Change
'
'End Sub

Private Sub txtSearch_Change()

    If Len(txtSearch) = 0 Then
        lvFiltered.Visible = False
        Exit Sub
    End If
    
    Dim li As ListItem
    Dim li2 As ListItem
    Dim llv As ListView
    
    'If SSTab1.Tab = 0 Then Set llv = lv Else Set llv = lv2
    Set llv = lv2
    
    lvFiltered.Visible = True
    lvFiltered.ListItems.Clear
    
    For Each li In llv.ListItems
        If txtSearch = "checked" And li.Checked Then
            Set li2 = lvFiltered.ListItems.Add(, , li.text)
            Set li2.Tag = li.Tag
            li2.Checked = li.Checked
        ElseIf InStr(1, li.text, txtSearch, vbTextCompare) > 0 Then
            Set li2 = lvFiltered.ListItems.Add(, , li.text)
            Set li2.Tag = li.Tag
            li2.Checked = li.Checked
        End If
    Next
        
        
End Sub

Private Function RefAlreadyExists(clsid As String) As Boolean
    Dim r As Reference
    For Each r In VBInstance.ActiveVBProject.References
        If InStr(1, r.guid, clsid, vbTextCompare) > 0 Then
            RefAlreadyExists = True
            Exit Function
        End If
    Next
End Function

Private Function GetReference(clsid As String, Optional isName As Boolean = False) As Reference
    Dim r As Reference
    For Each r In VBInstance.ActiveVBProject.References
        Debug.Print r.guid & " : " & r.name
        If isName Then
            If InStr(1, clsid, r.name, vbTextCompare) > 0 Then
                Set GetReference = r
                Exit Function
            End If
        Else
            If InStr(1, r.guid, clsid, vbTextCompare) > 0 Then
                Set GetReference = r
                Exit Function
            End If
        End If
    Next
End Function

