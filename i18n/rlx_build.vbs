' rlx_build.vbs  -  Excel automation helper for the RelaxTools English build.
'
' !!  NOT TESTED  !!  This script was written without access to Excel.  Always run it on COPIES and
' check the log.  It needs Windows + desktop Excel and
'   File > Options > Trust Center > Trust Center Settings > Macro Settings >
'   "Trust access to the VBA project object model"   (tick it, close all Excel windows first)
'
' Usage (from a command prompt in the i18n folder):
'   cscript //nologo rlx_build.vbs import  <original RelaxTools.xlam> <new.xlam> <build_en\Source\src>
'       copy the original xlam, replace all standard/class/form modules with the translated ones and
'       replace the code of ThisWorkbook / Sheet*  (cell contents of the sheets are NOT touched)
'   cscript //nologo rlx_build.vbs dump    <xlam> <frx_dump.txt>
'       write every Caption / ControlTipText / Text of every UserForm control (UTF-16 TSV)
'       -> then run:  python rlx_i18n.py import-frx frx_dump.txt
'   cscript //nologo rlx_build.vbs apply-forms <xlam> <build_en\frx_translations.txt>
'       set translated Caption / ControlTipText / Text on every UserForm control whose current
'       text exactly equals a Japanese source string, then save
'
' Afterwards put the English ribbon in:  python rlx_i18n.py patch-xlam new.xlam RelaxTools_en.xlam
Option Explicit

Dim fso, args, xl, logCount
Set fso = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments
logCount = 0

If args.Count < 1 Then Usage
Select Case LCase(args(0))
    Case "import"
        If args.Count < 4 Then Usage
        DoImport args(1), args(2), args(3)
    Case "dump"
        If args.Count < 3 Then Usage
        DoDump args(1), args(2)
    Case "apply-forms"
        If args.Count < 3 Then Usage
        DoApplyForms args(1), args(2)
    Case Else
        Usage
End Select
WScript.Quit 0

Sub Usage()
    WScript.Echo "usage: cscript //nologo rlx_build.vbs import|dump|apply-forms ...   (see header of this file)"
    WScript.Quit 1
End Sub

Sub Log(msg)
    WScript.Echo msg
    logCount = logCount + 1
End Sub

Function NewExcel()
    Dim x
    Set x = CreateObject("Excel.Application")
    x.Visible = False
    x.DisplayAlerts = False
    x.EnableEvents = False
    x.AutomationSecurity = 3          ' msoAutomationSecurityForceDisable: do not run the add-in's own startup code
    Set NewExcel = x
End Function

Function OpenProject(path, ByRef wb)
    Set xl = NewExcel()
    Set wb = xl.Workbooks.Open(fso.GetAbsolutePathName(path))
    On Error Resume Next
    Dim p
    Set p = wb.VBProject
    If Err.Number <> 0 Then
        Log "Cannot access the VBA project. Enable 'Trust access to the VBA project object model' in the Trust Center."
        wb.Close False: xl.Quit
        WScript.Quit 2
    End If
    On Error GoTo 0
    Set OpenProject = p
End Function

Sub CloseAll(wb, doSave)
    If doSave Then wb.Save
    wb.Close False
    xl.Quit
End Sub

' ------------------------------------------------------------------ import
Sub DoImport(container, outPath, srcDir)
    Dim wb, vbp, i, comp, sd, folder, f, ext, n
    fso.CopyFile container, outPath, True
    Set vbp = OpenProject(outPath, wb)

    For i = vbp.VBComponents.Count To 1 Step -1
        Set comp = vbp.VBComponents(i)
        If comp.Type = 1 Or comp.Type = 2 Or comp.Type = 3 Then vbp.VBComponents.Remove comp
    Next

    n = 0
    For Each sd In Array("Modules", "Class", "Form")
        If fso.FolderExists(srcDir & "\" & sd) Then
            For Each f In fso.GetFolder(srcDir & "\" & sd).Files
                ext = LCase(fso.GetExtensionName(f.Name))
                If ext = "bas" Or ext = "cls" Or ext = "frm" Then
                    On Error Resume Next
                    vbp.VBComponents.Import f.Path
                    If Err.Number <> 0 Then
                        Log "IMPORT FAILED: " & f.Name & " : " & Err.Description
                        Err.Clear
                    Else
                        n = n + 1
                    End If
                    On Error GoTo 0
                End If
            Next
        End If
    Next
    Log "imported " & n & " modules/classes/forms"

    n = 0
    sd = srcDir & "\Microsoft Excel Objects"
    If fso.FolderExists(sd) Then
        For Each f In fso.GetFolder(sd).Files
            If LCase(fso.GetExtensionName(f.Name)) = "cls" Then
                If ReplaceDocCode(vbp, f.Path) Then n = n + 1
            End If
        Next
    End If
    Log "replaced code of " & n & " document modules (ThisWorkbook / Sheet*)"
    CloseAll wb, True
    Log "saved " & outPath & "   (" & logCount & " log lines; look for FAILED / not found)"
End Sub

Function ReplaceDocCode(vbp, path)
    Dim ts, lines, i, l, nm, inBegin, started, code, cm, re, m
    ReplaceDocCode = False
    Set ts = fso.OpenTextFile(path, 1, False, 0)
    lines = Split(ts.ReadAll, vbLf)
    ts.Close
    Set re = New RegExp
    re.Pattern = "^Attribute VB_Name = ""(\w+)"""
    nm = "": inBegin = False: started = False: code = ""
    For i = 0 To UBound(lines)
        l = lines(i)
        If Right(l, 1) = vbCr Then l = Left(l, Len(l) - 1)
        If Not started Then
            If re.Test(l) Then
                nm = re.Execute(l)(0).SubMatches(0)
            ElseIf Left(l, 8) = "VERSION " Or Left(l, 10) = "Attribute " Then
                ' header
            ElseIf l = "BEGIN" Then
                inBegin = True
            ElseIf inBegin Then
                If l = "END" Then inBegin = False
            Else
                started = True
            End If
        End If
        If started Then code = code & l & vbCrLf
    Next
    If nm = "" Then Log "no VB_Name in " & path: Exit Function
    On Error Resume Next
    Set cm = vbp.VBComponents(nm).CodeModule
    If Err.Number <> 0 Then
        Log "document module not found in xlam: " & nm
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString code
    ReplaceDocCode = True
End Function

' ------------------------------------------------------------------ forms: dump / apply
Function Esc(s)
    s = Replace(s, "\", "\\")
    s = Replace(s, vbTab, "\t")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    Esc = s
End Function

Function Unesc(s)
    Dim i, c, nx, r
    r = "": i = 1
    Do While i <= Len(s)
        c = Mid(s, i, 1)
        If c = "\" And i < Len(s) Then
            nx = Mid(s, i + 1, 1)
            If nx = "t" Then
                r = r & vbTab
            ElseIf nx = "r" Then
                r = r & vbCr
            ElseIf nx = "n" Then
                r = r & vbLf
            Else
                r = r & nx
            End If
            i = i + 2
        Else
            r = r & c
            i = i + 1
        End If
    Loop
    Unesc = r
End Function

Sub DoDump(xlamPath, outPath)
    Dim wb, vbp, comp, d, ctl, out, seen
    Set vbp = OpenProject(xlamPath, wb)
    Set out = fso.CreateTextFile(outPath, True, True)      ' UTF-16
    Set seen = CreateObject("Scripting.Dictionary")
    For Each comp In vbp.VBComponents
        If comp.Type = 3 Then
            Set d = Nothing
            On Error Resume Next
            Set d = comp.Designer
            On Error GoTo 0
            If d Is Nothing Then
                Log "no designer for " & comp.Name
            Else
                EmitText out, seen, comp.Name, "(form)", "Caption", ReadProp(d, "Caption")
                For Each ctl In d.Controls
                    WalkControl out, seen, comp.Name, ctl, False, Nothing
                Next
            End If
        End If
    Next
    out.Close
    Log "dumped " & seen.Count & " control strings to " & outPath
    CloseAll wb, False
End Sub

Function ReadProp(obj, prop)
    Dim v
    v = Empty
    On Error Resume Next
    Select Case prop
        Case "Caption": v = obj.Caption
        Case "ControlTipText": v = obj.ControlTipText
        Case "Text": v = obj.Text
    End Select
    On Error GoTo 0
    ReadProp = v
End Function

Sub EmitText(out, seen, frm, ctlName, prop, v)
    Dim k
    If VarType(v) = 8 Then
        If Len(v) > 0 Then
            k = frm & "|" & ctlName & "|" & prop
            If Not seen.Exists(k) Then
                seen.Add k, 1
                out.WriteLine frm & vbTab & ctlName & vbTab & prop & vbTab & Esc(v)
            End If
        End If
    End If
End Sub

' mode False = dump ; apply is done by ApplyControl (below)
Sub WalkControl(out, seen, frm, ctl, applyMode, dict)
    Dim nm, tn, i, p
    nm = "": tn = ""
    On Error Resume Next
    nm = ctl.Name
    tn = TypeName(ctl)
    On Error GoTo 0
    EmitText out, seen, frm, nm, "Caption", ReadProp(ctl, "Caption")
    EmitText out, seen, frm, nm, "ControlTipText", ReadProp(ctl, "ControlTipText")
    EmitText out, seen, frm, nm, "Text", ReadProp(ctl, "Text")
    On Error Resume Next
    If tn = "MultiPage" Then
        For i = 0 To ctl.Pages.Count - 1
            EmitText out, seen, frm, nm & ".Pages(" & i & ")", "Caption", ctl.Pages(i).Caption
            EmitText out, seen, frm, nm & ".Pages(" & i & ")", "ControlTipText", ctl.Pages(i).ControlTipText
        Next
    ElseIf tn = "TabStrip" Then
        For i = 0 To ctl.Tabs.Count - 1
            EmitText out, seen, frm, nm & ".Tabs(" & i & ")", "Caption", ctl.Tabs(i).Caption
            EmitText out, seen, frm, nm & ".Tabs(" & i & ")", "ControlTipText", ctl.Tabs(i).ControlTipText
        Next
    End If
    On Error GoTo 0
End Sub

Function LoadDict(path)
    Dim ts, lines, l, parts, d, pos
    Set d = CreateObject("Scripting.Dictionary")
    Set ts = fso.OpenTextFile(path, 1, False, -1)          ' Unicode
    Do While Not ts.AtEndOfStream
        l = ts.ReadLine
        pos = InStr(l, vbTab)
        If pos > 1 Then
            If Not d.Exists(Unesc(Left(l, pos - 1))) Then d.Add Unesc(Left(l, pos - 1)), Unesc(Mid(l, pos + 1))
        End If
    Loop
    ts.Close
    Set LoadDict = d
End Function

Sub TrySet(obj, prop, dict, ByRef n)
    Dim v
    v = ReadProp(obj, prop)
    If VarType(v) = 8 Then
        If dict.Exists(v) Then
            On Error Resume Next
            Select Case prop
                Case "Caption": obj.Caption = dict(v)
                Case "ControlTipText": obj.ControlTipText = dict(v)
                Case "Text": obj.Text = dict(v)
            End Select
            If Err.Number <> 0 Then
                Log "could not set " & prop & " : " & Err.Description
                Err.Clear
            Else
                n = n + 1
            End If
            On Error GoTo 0
        End If
    End If
End Sub

Sub DoApplyForms(xlamPath, transPath)
    Dim wb, vbp, comp, d, ctl, dict, n, i, tn
    Set dict = LoadDict(transPath)
    Log "loaded " & dict.Count & " translations"
    Set vbp = OpenProject(xlamPath, wb)
    n = 0
    For Each comp In vbp.VBComponents
        If comp.Type = 3 Then
            Set d = Nothing
            On Error Resume Next
            Set d = comp.Designer
            On Error GoTo 0
            If Not d Is Nothing Then
                TrySet d, "Caption", dict, n
                For Each ctl In d.Controls
                    TrySet ctl, "Caption", dict, n
                    TrySet ctl, "ControlTipText", dict, n
                    TrySet ctl, "Text", dict, n
                    tn = ""
                    On Error Resume Next
                    tn = TypeName(ctl)
                    If tn = "MultiPage" Then
                        For i = 0 To ctl.Pages.Count - 1
                            TrySet ctl.Pages(i), "Caption", dict, n
                            TrySet ctl.Pages(i), "ControlTipText", dict, n
                        Next
                    ElseIf tn = "TabStrip" Then
                        For i = 0 To ctl.Tabs.Count - 1
                            TrySet ctl.Tabs(i), "Caption", dict, n
                            TrySet ctl.Tabs(i), "ControlTipText", dict, n
                        Next
                    End If
                    On Error GoTo 0
                Next
            End If
        End If
    Next
    Log "applied " & n & " control texts"
    CloseAll wb, True
End Sub
