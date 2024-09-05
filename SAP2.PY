import win32com.client

If not IsObject(application) Then
    Set SapGuiAuto = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
End if 
If Not IsObject(connection) Then
    Set connection = application.Children(0)
End if 
If Not IsObject(session) Then
    Set session = connection.Children(0)
End if 
If IsObject(WScript) Then
    WScript.connectObject session, "on"
    WScript.connectObject application, "on"
End if 

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

session.findById("wnd[0]").resizeWorkingPane 117,22,false
