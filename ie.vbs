    Const HKEY_CURRENT_USER = &H80000001
    Dim strComputer 

    strComputer = "."
    Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
      
    strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
     
    strValueName = "ProxyEnable"
    dwValue = 1
    objRegistry.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue
     
    strValueName = "ProxyServer"
    strValue = "192.168.1.1:8080"
    objRegistry.SetStringValue HKEY_CURRENT_USER, strKeyPath, strValueName, strValue
     
    strValueName = "ProxyOverride"
    strValue = "<local>"
    objRegistry.SetStringValue HKEY_CURRENT_USER, strKeyPath, strValueName, strValue

