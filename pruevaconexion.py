import win32com.client as win32

def conexion (): 
   
    SapGuiAuto = win32.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32.CDispatch:
        return
    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32.CDispatch:
        SapGuiAuto = None
        return
    connection = application.Children(0)
    if not type(connection) == win32.CDispatch:
        application = None
        SapGuiAuto = None
        return
    session = connection.Children(0)
    if not type(session) == win32.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return
    return session

def ingresar_transaccion():

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "573891"
    session.findById("wnd[0]/usr/ctxtRMMG1-AENNR").setFocus()
    session.findById("wnd[0]/usr/ctxtRMMG1-AENNR").caretPosition = 0

if __name__ == "__main__":
    global session
    session = conexion()
    ingresar_transaccion()

