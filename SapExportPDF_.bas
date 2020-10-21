Attribute VB_Name = "SapExportPDF_"
Sub SapExportPDF()

Set SapGuiAuto = GetObject("SAPGUI")
Set SapApp = SapGuiAuto.GetScriptingEngine
Set Connection = SapApp.Children(0)
Set session = Connection.Children(0)
Dim StartTime As Double
Dim MinutesElapsed As String

StartTime = Timer
session.FindById("wnd[0]").Maximize

i = 2
Z = 2
SavePath = Cells.Range("C1").Value

While Cells.Range("a" & i).Value <> ""

    session.FindById("wnd[0]").ResizeWorkingPane 170, 39, False
    file_name = Cells.Range("A" & i).Value
    session.FindById("wnd[0]/usr/ctxtVBRK-VBELN").Text = file_name
    session.FindById("wnd[0]/mbar/menu[0]/menu[11]").Select
    
    session.FindById("wnd[1]/tbar[0]/btn[6]").Press
    
    session.FindById("wnd[2]/usr/chkNAST-DIMME").Selected = True
    session.FindById("wnd[2]/usr/chkNAST-DELET").Selected = True
    session.FindById("wnd[2]/usr/chkNAST-DELET").SetFocus
    session.FindById("wnd[2]/tbar[0]/btn[0]").Press
    session.FindById("wnd[1]/tbar[0]/btn[86]").Press
    Application.Wait Now + TimeValue("0:00:03")

    Set WshShell = CreateObject("WScript.Shell")
    
        file_name = Cells.Range("A" & i) ' the name of the invoiced searched in SAP
        
        WshShell.SendKeys "{tab}"
        Application.Wait Now + TimeValue("0:00:01")
        WshShell.AppActivate "Save As" ' search for Save As box
        
        'WshShell.SendKeys "^+s"
        Application.Wait Now + TimeValue("0:00:01") ' wait to be sure windows is created
        WshShell.SendKeys "%n" ' go to insert name
        WshShell.SendKeys SavePath & "\" & file_name & ".pdf" ' path and name
        Application.Wait Now + TimeValue("0:00:01")
        WshShell.SendKeys "%s" ' save activity
         ' wait to be sure everything is processed
    
    Set WshShell = Nothing
    
    i = i + 1

Wend

MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
MsgBox " minute " & MinutesElapsed

End Sub


