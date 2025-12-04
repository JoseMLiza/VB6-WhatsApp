Attribute VB_Name = "mdlAplication"
'=========================================================================
'
' Project   : VB6-WhatsApp
' Module    : mdlAplication.bas
' Author    : Jose Liza (https://github.com/JoseMLiza)
'
'=========================================================================

Option Explicit

'***********************
'* PUBLIC DECLARATIONS *
'***********************
Public Enum enmWebApp
    EdgeWebView2
    SeleniumBasic
End Enum

Public Enum enmSEWebDriver
    None
    EdgeDriver
    ChromeDriver
End Enum

Public Const UpdaterName            As String = "WebDriver-Update" ' WebDriver-Update.exe
Public Const SeleniumBase           As String = "\bin\SeleniumBasic\"
Public Const SeEdgeDriverPath       As String = SeleniumBase & "edgedriver.exe"
Public Const SeChromeDriverPath     As String = SeleniumBase & "chromedriver.exe"
Public Const tempDownload           As String = SeleniumBase & "download"

Public Const UserAgentChrome        As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{VERSION} Safari/537.36"
Public Const UserAgentMsEdge        As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36 Edg/{VERSION}"

Public Const WebSiteMSEdge          As String = "https://developer.microsoft.com/es-es/microsoft-edge/tools/webdriver?form=MA13LH"
Public Const WebSiteChrom           As String = "https://googlechromelabs.github.io/chrome-for-testing/"
Public Const BaseLinkMSEdge         As String = "https://msedgedriver.microsoft.com/"                       '125.0.2535.51/edgedriver_win32.zip"
Public Const BaseLinkChrome         As String = "https://storage.googleapis.com/chrome-for-testing-public/" '125.0.6422.60/win32/chromedriver-win32.zip"

Public m_ExitCodeUpdate     As Integer
Public m_IsCompiled         As Boolean
Public m_IsHeadless         As Boolean
Public m_HwndWebBrowser     As Long

Public m_wbVersion          As String
Public m_wdVersion          As String

Public webApp               As enmWebApp
Public SEWebDriver          As enmSEWebDriver

Public New_c                As cConstructor

Public m_UserDataDir        As String

'Public objSEWebDriver       As Object 'Compiled
Public objSEWebDriver       As Selenium.WebDriver 'Development
Public objSEBy              As Selenium.bY
Public objKeys              As Selenium.Keys

Public arrDataEmojis(4, 2)  As String 'Diccionario de emojis (se recomienda usar base de datos).

'************************
'* PRIVATE DECLARATIONS *
'************************
'Kernel32.dll
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreatePipe Lib "kernel32.dll" (ByRef hReadPipe As Long, ByRef hWritePipe As Long, ByRef lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32.dll" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Any, ByVal lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Any, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function LoadLibraryW Lib "kernel32.dll" (ByVal lpLibFileName As Long) As Long
Private Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function SetHandleInformation Lib "kernel32.dll" (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'DirectCOM.dll (RC6)
Private Declare Function GetInstanceEx Lib "DirectCOM" (spFName As Long, spClassName As Long, Optional ByVal UseAlteredSearchPath As Boolean = True) As Object

'Shell32.dll
'Private Declare Function IsUserAnAdmin Lib "Shell32.dll" () As Long
'Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteEx Lib "Shell32.dll" Alias "ShellExecuteExA" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long

'Advapi32.dll
Private Declare Function RegCloseKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "Advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "Advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

'User define type
Private Type POINT
    X As Long
    Y As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

'#If Win64 Then
'    Private Const KEY_WOW64_32KEY   As Long = &H200&
'    Private Const KEY_WOW64_64KEY   As Long = &H100&
'#Else
'    Private Const KEY_WOW64_32KEY   As Long = &H200&
'    Private Const KEY_WOW64_64KEY   As Long = &H0&
'#End If

'Para ShellExecute o ShellExecuteEx (ocultar o mostrar ventana)
Private Const ERROR_CANCELLED           As Long = &H4C7 '1223
Private Const HANDLE_FLAG_INHERIT       As Long = &H1
Private Const INFINITE                  As Long = &HFFFF
Private Const SEE_MASK_NOCLOSEPROCESS   As Long = &H40
Private Const SEE_MASK_NO_CONSOLE       As Long = &H8000000
Private Const SEE_MASK_FLAG_NO_UI       As Long = &H400
Private Const STARTF_USESTDHANDLES      As Long = &H100
Private Const SW_HIDE                   As Long = 0
Private Const SW_NORMAL                 As Long = 1

'Para validar registro del sistema
Private Const HKEY_CLASSES_ROOT         As Long = &H80000000
Private Const HKEY_CURRENT_USER         As Long = &H80000001
Private Const KEY_ALL_ACCESS            As Long = &H2003F
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const REG_SZ                    As Long = 1
Private Const ERROR_SUCCESS             As Long = 0

Private Const RegAsmPath                As String = "\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
Private Const DirectComDllRelPath       As String = "\bin\RC6\DirectCOM.dll"
Private Const RCDllRelPath              As String = "\bin\RC6\RC6.dll"
Private Const SeleniumDll               As String = SeleniumBase & "Selenium.dll"
Private Const SeleniumTlb               As String = SeleniumBase & "Selenium.tlb"
Private Const EdgeNameDriverDownload    As String = "edgedriver_win32.zip"
Private Const ChromeNameDriverDownload  As String = "chromedriver-win32.zip"

Private m_RegAsmPath        As String

'********************
'* PUBLIC FUNCTIONS *
'********************
Public Function WebAppConstruct(ByVal pWebApp As enmWebApp) As Boolean
    On Error GoTo FunctionError
    '---
    Dim lResult As Long
    '---
    WebAppConstruct = False
    '---
    Select Case pWebApp
        Case EdgeWebView2
            'MICROSOFT EDGE WEBVIEW2
            Static st_RC As cConstructor
            If Not st_RC Is Nothing Then Set New_c = st_RC: Exit Function
            '---
            'If App.LogMode Then  'we run compiled - and try to ensure regfree instantiation from \Bin\
                On Error Resume Next
                    LoadLibraryW StrPtr(App.Path & DirectComDllRelPath)
                    Set st_RC = GetInstanceEx(StrPtr(App.Path & RCDllRelPath), StrPtr("cConstructor"))
                    If st_RC Is Nothing Then MsgBox "Couldn't load regfree (RC6), will try with a registered version next...", vbInformation + vbOKOnly
                On Error GoTo 0
            'End If
            If st_RC Is Nothing Then Set st_RC = cGlobal.New_c 'fall back to loading a registered version
            If st_RC Is Nothing Then Exit Function
            Set New_c = st_RC
        Case SeleniumBasic
            'SELENIUMBASIC
            '**VALIDACION DE SELENIUM**
            ' * - Intentara setear la variable de tipo objeto *objSEWebDriver*, con CreateObject desde la funcion *SeCreateObject*:
            '     - True : Procederá a trabajar con selenium
            '     - False: Validará si existe en el registro del sistema, la libreria *Selenium.dll*
            '              Si no existe en el registro, lo registrará usando *RegAsm*, ya que la libreria esta basada en .NET, con este proceso generaremos el TLB de *Selenium*
            ' * -
            If Not SeCreateObject Then
                MsgBox "Unable to create the object required for this method.", vbExclamation + vbOKOnly
                '--> Continua si es administrador
                If Not SeIsDLLRegistered(lResult) Then
                    If MsgBox("SE(Selenium) is not registered on your system." & vbCrLf & vbCrLf & "Do you want to register it?", vbQuestion + vbYesNo) = vbNo Then Exit Function
                    '--> Continua si se acepta registrarlo.
                    'Si la validacion al registro devuelve valor 8: entonces no tiene acceso al registro del sistema.
                    If lResult = 8 Then
                        MsgBox "You do not have access to the registry, check with your system administrator", vbCritical + vbOKOnly, "Access denied"
                        Exit Function
                    End If
                    'Proceder a registrar Selenium.dll
                    If Not RegisterAsmDLL(App.Path & SeleniumDll, App.Path & SeleniumTlb) Then Exit Function
                End If
            End If
            '---
            'Set objSEWebDriver = CreateObject("Selenium.EdgeDriver")
            Select Case SEWebDriver
                Case None
                    Set objSEWebDriver = Nothing
                Case EdgeDriver
                    Set objSEWebDriver = New Selenium.EdgeDriver
                Case ChromeDriver
                    Set objSEWebDriver = New Selenium.ChromeDriver
            End Select
            Set objSEBy = New Selenium.bY
            Set objKeys = New Selenium.Keys
            '---
    End Select
    '---
    WebAppConstruct = True
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function UpdateWebDriver(ByVal webBrowser As enmSEWebDriver, ByRef pathDownloadFile As String, ByRef error As Boolean) As Boolean
    Dim strLink As String, exePath As String, argParameters As String
    Dim bUpdate As Boolean, result As Boolean
    
    error = False
    result = False
    strLink = GetLinkWebDriver(SEWebDriver, m_wbVersion, m_wdVersion)
    '---
    'VERSION DEL WEBBROWSER
    If m_wbVersion = vbNullString Then
        error = True
        MsgBox "You do not have the selected browser installed, contact your system administrator.", vbExclamation + vbOKOnly
        GoTo Continue
    End If
    '---
    'VERSION DEL WEBDRIVER
    If m_wdVersion = vbNullString Then
        MsgBox "The webdriver necessary for the application to function has not been found, it will start with the download." & vbCrLf & vbCrLf & _
               "WebDriver: " & IIf(SEWebDriver = EdgeDriver, "Microsoft Edge", "Google Chrome") & vbCrLf & _
               "Version: " & Trim(m_wbVersion), vbInformation + vbOKOnly
        bUpdate = True
        GoTo Continue
    End If
    '---
    If Not strLink = vbNullString Then
        If MsgBox("It has been detected that the version of your WebDriver is different from your selected browser." & vbCrLf & vbCrLf & _
                  "To ensure the correct functioning of the application, it is necessary to update the WebDriver to the version of your browser." & vbCrLf & vbCrLf & _
                  "Do you want to continue with the update?" & vbCrLf & vbCrLf & _
                  "Current version of WebDriver : " & Trim(m_wdVersion) & vbCrLf & _
                  "Current version of WebBrowser: " & Trim(m_wbVersion), vbQuestion + vbYesNo, "") = vbYes Then
            bUpdate = True
        End If
    End If
    '---
Continue:
    If bUpdate Then
        exePath = App.Path & "\" & UpdaterName & ".exe"
        argParameters = "--title ""WebDriver update"" --url """ & strLink & """  --path """ & App.Path & tempDownload & """"
        '---
        result = RunAsExe(exePath, argParameters, False)
        If result Then
            pathDownloadFile = App.Path & tempDownload & "\" & IIf(webBrowser = EdgeDriver, EdgeNameDriverDownload, ChromeNameDriverDownload)
            pathDownloadFile = IIf(FileExists(pathDownloadFile), pathDownloadFile, "")
        End If
        '---
    End If
    UpdateWebDriver = result
End Function

Public Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

Public Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) <> 0
    On Error GoTo 0
End Function

'Public Function GetLinkWebDriver(ByVal webBrowser As enmSEWebDriver, Optional ByVal html As HTMLDocument) As String
Public Function GetLinkWebDriver(ByVal webBrowser As enmSEWebDriver, Optional ByRef strWBVersion As String, Optional ByRef strWDVersion As String) As String
    On Error GoTo FunctionError
    Dim strLink As String
    Dim pbWebSite As Boolean
    '---
    pbWebSite = False
    strWBVersion = GetVersionWebBrowser(webBrowser)
    If strWBVersion = vbNullString Then Exit Function
    '---
    strWDVersion = GetVersionWebDriver(webBrowser)
    '---
    If strWBVersion <> strWDVersion Then
        Select Case webBrowser
            Case EdgeDriver
                strLink = BaseLinkMSEdge & strWBVersion & "/" & EdgeNameDriverDownload '"/edgedriver_win32.zip"
            Case ChromeDriver
                strLink = BaseLinkChrome & strWBVersion & "/win32/" & ChromeNameDriverDownload '"/win32/chromedriver-win32.zip"
        End Select
    End If
    GetLinkWebDriver = strLink
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function ReadFileExitCode(ByVal filePath As String) As Integer
    On Error GoTo FunctionError
    '---
    Dim n As Integer
    Dim exitCode As String
    n = FreeFile
    Open filePath For Input As #n
    Line Input #n, exitCode
    Close #n
    ReadFileExitCode = CLng(exitCode)
    Exit Function
    '---
FunctionError:
    ReadFileExitCode = -1 ' Indica que no se pudo leer
    Call MessageError(Err)
End Function

Public Function ReplaceEmojiInText(ByVal strMessage As String) As String
    Dim strPalabras() As String, strOutPut As String
    Dim i, j As Long
    strPalabras = Split(strMessage, Space(1))
    For i = 0 To UBound(strPalabras)
        If InStr(strPalabras(i), "[emoji_") > 0 Then
            For j = 0 To UBound(arrDataEmojis)
                If strPalabras(i) = arrDataEmojis(j, 0) Or strPalabras(i) = arrDataEmojis(j, 1) Then
                    strPalabras(i) = GetStrEmojiFromUnicode(arrDataEmojis(j, 2))
                End If
            Next
        End If
        strOutPut = strOutPut + strPalabras(i) & IIf(i < UBound(strPalabras), Space(1), "")
        ReplaceEmojiInText = strOutPut
    Next
End Function

Public Function GetStrEmojiFromUnicode(ByVal emojiCode As String) As String
    Dim lngCode As Long
    Dim highCode As Integer
    Dim lowCode As Integer
    Dim output As String
    '---
    If Left$(emojiCode, 2) = "U+" Then emojiCode = Mid(emojiCode, 3)
    '---
    lngCode = CLng("&H" & emojiCode)
    If lngCode <= &HFFFF Then
        output = ChrW(emojiCode)
    Else
        lngCode = lngCode - &H10000
        highCode = &HD800 Or ((lngCode \ &H400) And &H3FF)
        lowCode = &HDC00 Or (lngCode And &H3FF)
        output = ChrW(highCode) & ChrW(lowCode)
    End If
    '---
    GetStrEmojiFromUnicode = output
End Function

'*********************
'* PRIVATE FUNCTIONS *
'*********************
'Private Function GetLinkFromWebSite(ByVal html As HTMLDocument, ByVal webBrowser As enmSEWebDriver) As String
'    On Error GoTo FunctionError
'    '---
'    Dim objElements As Object, objElement() As Object
'    Dim result As String
'    Dim i As Integer
'    Select Case webBrowser
'        Case EdgeDriver
'            '*************************
'            '* OBTENER DESDE WEBSITE *
'            '*************************
'            'Redimencionamos para manejar 2 elementos:
'            ' 1. Elemento inicial <div> de clase 'common-pager__page--active'
'            ' 2. Elemento <span> para ubicar el enlace.
'            ReDim objElement(1) As Object
'            Set objElement(0) = html.querySelector("div.common-pager__page--active")
'            If Not objElement(0) Is Nothing Then
'                html.Body.innerHTML = objElement(0).innerHTML
'                Set objElements = html.querySelectorAll("a.common-button.common-button--loading.common-button--tag")
'                '---
'                For i = 0 To elementos.Length - 1
'                    'Buscar el primer <SPAN> dentro del elemento <A>
'                    Set objElement(1) = objElements.Item(i).GetElementsByTagName("SPAN").Item(0)
'                    'Verificar si el texto del <SPAN> es "x86"
'                    If Not objElement(1) Is Nothing And Trim(objElement(1).innerText) = "x86" Then
'                        'Obtener el enlace (href) del elemento <A>
'                        result = objElements.Item(i).getAttribute("href")
'                        Exit For
'                    End If
'                Next i
'                '---
'            End If
'            '***************************
'            '* CONCATENAR CON VARIABLE *
'            '***************************
'        Case ChromeDriver
'    End Select
'    '---
'    GetLinkFromWebSite = result
'    '---
'    Exit Function
'FunctionError:
'    Call MessageError(Err)
'End Function
Private Function GetFileNameFromPath(ByVal filePath As String) As String
    GetFileNameFromPath = Mid$(filePath, InStrRev(filePath, "\") + 1)
End Function

Private Function RegisterAsmDLL(ByVal dllPath As String, ByVal tlbPath As String) As Boolean
    Dim exePath As String, argParameters As String
    '---
    RegisterAsmDLL = False
    '---
    If CheckFileRegAsm Then
        exePath = m_RegAsmPath
        argParameters = """" & dllPath & """ /codebase /tlb:" & """" & tlbPath & """"
        '---
        If RunAsExe(exePath, argParameters, False) Then
            MsgBox "SE (Selenium) was registered successfully", vbInformation + vbOKOnly, "Registered successfully"
        Else
            MsgBox "An error occurred while registering the SE library (Selenium), contact your system administrator", vbCritical + vbOKOnly, "Registration error"
            Exit Function
        End If
        '---
    Else
        MsgBox "Your system does not have the .NET Framework 4, this is necessary for the SE (Selenium) registration.", vbCritical + vbOKOnly, ".NET Framework 4 error"
        Exit Function
    End If
    RegisterAsmDLL = True
End Function

Private Function GetSystemDrive() As String
    Dim winDir      As String
    Dim systemDrive As String
    
    ' Obtener el directorio de Windows
    winDir = String(255, 0)
    GetWindowsDirectory winDir, Len(winDir)
    systemDrive = Left(winDir, 2)
    
    GetSystemDrive = systemDrive
End Function

Private Function CheckFileRegAsm() As Boolean
    Dim systemDrive As String
    
    CheckFileRegAsm = False
    
    ' Obtener la unidad en la que está instalado el sistema operativo
    systemDrive = GetSystemDrive()
    
    ' Archivo a buscar
    m_RegAsmPath = systemDrive & RegAsmPath
    
    ' Verificar si el archivo existe en la unidad del sistema
    CheckFileRegAsm = FileExists(m_RegAsmPath)
    
    CheckFileRegAsm = True
End Function

Private Function SeIsDLLRegistered(Optional ByRef lState As Long) As Boolean
    Dim seGUID  As String
    Dim hKey    As Long
    Dim lResult As Long
    
    SeIsDLLRegistered = False
    
    'GUID de ("Selenium.Application") = {0277FC34-FD1B-4616-BB19-E9AAFA695FFB}
    'GUID de ("Selenium.WebDriver") = {0277FC34-FD1B-4616-BB19-E3CCFFAB4234}
    seGUID = "{0277FC34-FD1B-4616-BB19-E3CCFFAB4234}"

    ' Componer la ruta de la clave CLSID
    Dim regKeyCLSID As String
    regKeyCLSID = "CLSID\" & seGUID & "\InprocServer32"

    ' Abrir la clave en HKCR
    'lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, regKeyCLSID, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY Or KEY_WOW64_32KEY, hKey)
    lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, regKeyCLSID, 0, KEY_ALL_ACCESS, hKey)
    If lResult = 0 Then
        ' La clave está presente, la DLL está registrada
        SeIsDLLRegistered = True
        ' Cerrar la clave
        RegCloseKey hKey
    End If
    'Devolver el codigo de resultado
    lState = lResult
End Function

Private Function GetVersionWebBrowser(ByVal webBrowser As enmSEWebDriver) As String
    On Error GoTo FunctionError
    Dim hKey As Long, lResult As Long, lDataSize As Long, lpType As Long
    Dim lpData As String
    lDataSize = 1024
    '---
    'Abrir la clave del registro
    Select Case webBrowser
        Case EdgeDriver
            lResult = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Edge\BLBeacon", 0, KEY_QUERY_VALUE, hKey)
        Case ChromeDriver
            lResult = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Google\Chrome\BLBeacon", 0, KEY_QUERY_VALUE, hKey)
    End Select
    If lResult <> ERROR_SUCCESS Then
        Err.Raise 5100, App.EXEName & ".GetVersionWebBrowser", "Error opening the registry key or the web browser is not installed."
    End If
    
    'Obtener el tamaño del valor
    lResult = RegQueryValueEx(hKey, "version", 0, lpType, ByVal 0, lDataSize)
    If lResult <> ERROR_SUCCESS Then
        RegCloseKey hKey
        Err.Raise 5101, App.EXEName & ".GetVersionWebBrowser", "Error getting size of registry value."
    End If
    '---
    
    'Leer el valor
    lpData = Space$(lDataSize)
    lResult = RegQueryValueEx(hKey, "version", 0, lpType, lpData, lDataSize)
    If lResult <> ERROR_SUCCESS Then
        RegCloseKey hKey
        Err.Raise 5102, App.EXEName & ".GetVersionWebBrowser", "Error reading registry value."
    End If
    
    'Cerrar la clave del registro
    RegCloseKey hKey
    
    'Devolver el valor leído
    GetVersionWebBrowser = Left$(lpData, lDataSize - 1)
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Private Function GetVersionWebDriver(ByVal webBrowser As enmSEWebDriver) As String
    On Error GoTo FunctionError
    Dim command As String, params As String, output As String
    Dim varArrOut() As String
    Dim varArrVersion() As String
    Dim strVersion As String
    '---
    'Comando para obtener la versión del Microsoft Edge WebDriver
    Select Case webBrowser
        Case EdgeDriver
            If Not FileExists(App.Path & SeEdgeDriverPath) Then Exit Function
            command = "" & App.Path & SeEdgeDriverPath & ""
        Case ChromeDriver
            If Not FileExists(App.Path & SeChromeDriverPath) Then Exit Function
            command = "" & App.Path & SeChromeDriverPath & ""
    End Select
    params = "--version"
    
    'Ejecutar el comando y capturar la salida
    output = ExecuteCommand(command, params)
    
    'Devolver la versión del WebDriver
    Select Case webBrowser
        Case EdgeDriver
            varArrOut = Split(output, " ")
            strVersion = CStr(varArrOut(3))
            varArrVersion = Split(strVersion, ".")
        Case ChromeDriver
            varArrOut = Split(output, " ")
            strVersion = CStr(varArrOut(1))
            varArrVersion = Split(strVersion, ".")
    End Select
    
    GetVersionWebDriver = strVersion
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Private Function ExecuteCommand(ByVal command As String, ByVal params As String) As String
    Dim sa As SECURITY_ATTRIBUTES
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim hReadPipe As Long
    Dim hWritePipe As Long
    Dim buffer As String
    Dim bytesRead As Long
    Dim result As Long
    Dim output As String

    'Iniciar atributos de seguridad.
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1
    sa.lpSecurityDescriptor = 0

    'Crear pipe.
    result = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If result = 0 Then
        MsgBox "CreatePipe failed. Error: " & GetLastError()
        Exit Function
    End If

    'Asegúrese de que el identificador de lectura del pipe para STDOUT no se herede.
    result = SetHandleInformation(hReadPipe, HANDLE_FLAG_INHERIT, 0)
    If result = 0 Then
        MsgBox "SetHandleInformation failed. Error: " & GetLastError()
        Exit Function
    End If

    'Iniciar estructura STARTUPINFO.
    si.cb = Len(si)
    si.dwFlags = STARTF_USESTDHANDLES
    si.hStdOutput = hWritePipe
    si.hStdError = hWritePipe
    si.hStdInput = 0

    'Ejecutar el proceso
    result = CreateProcessA(vbNullString, command & " " & params, ByVal 0&, ByVal 0&, 1, 0, ByVal 0&, vbNullString, si, pi)
    If result = 0 Then
        MsgBox "CreateProcessA failed. Error: " & GetLastError()
        Exit Function
    End If

    'Cierre el extremo de escritura del pipe antes de leer.
    CloseHandle (hWritePipe)

    'Leer el resultado del proceso.
    Do
        buffer = Space(1024)
        result = ReadFile(hReadPipe, ByVal buffer, Len(buffer), bytesRead, ByVal 0&)
        If bytesRead > 0 Then
            output = output & Left(buffer, bytesRead)
        End If
    Loop While result <> 0 And bytesRead > 0

    'Esperar que el proceso se complete.
    WaitForSingleObject pi.hProcess, INFINITE

    'Cerrar objetos.
    CloseHandle (pi.hProcess)
    CloseHandle (pi.hThread)
    CloseHandle (hReadPipe)

    ExecuteCommand = output
End Function


Private Function SeCreateObject() As Boolean
    'Funcion para validar si Selenium esta registrado.
    On Error GoTo FunctionError
    Set objSEWebDriver = CreateObject("Selenium.WebDriver")
    SeCreateObject = True
    Exit Function
FunctionError:
    SeCreateObject = False
End Function

Private Function RunAsExe(ByVal exePath As String, ByVal params As String, Optional ByVal isShow As Boolean = False) As Boolean
    On Error GoTo RoutinError
    '---
    Dim ExecInfo    As SHELLEXECUTEINFO
    Dim success     As Long
    Dim lastError   As Long
    Dim exitCodePath As String
    '--
    'Configura la estructura SHELLEXECUTEINFO
    With ExecInfo
        .cbSize = Len(ExecInfo)
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .hwnd = 0
        .lpVerb = "runas"
        .lpFile = exePath
        .lpParameters = params
        .lpDirectory = vbNullString
        .nShow = IIf(isShow, SW_NORMAL, SW_HIDE)
        .hInstApp = 0
        .lpIDList = 0
        .lpClass = vbNullString
        .hkeyClass = 0
        .dwHotKey = 0
        .hIcon = 0
    End With
    
    'Ejecuta el proceso
    success = ShellExecuteEx(ExecInfo)
    
    If success = 0 Then
        lastError = Err.LastDllError
        
        'Verifica si el error es ERROR_CANCELLED
        If lastError = ERROR_CANCELLED Then
            MsgBox "The user canceled the elevation of permissions needed to execute the process.", vbExclamation, "Operación cancelada"
        Else
            Err.Raise success, App.EXEName & ".RunAsExe", "Error running the process, contact your system administrator."
        End If
        RunAsExe = False
        Exit Function
    End If
    
    'Espera a que el proceso termine
    WaitForSingleObject ExecInfo.hProcess, INFINITE
    
    'Cierra el identificador del proceso
    CloseHandle ExecInfo.hProcess
    
    'Leer archivo de resultado de la descarga
    exitCodePath = Environ$("TEMP") & "\" & UpdaterName & "_exitcode.update"
    m_ExitCodeUpdate = ReadFileExitCode(exitCodePath)
    Call Kill(exitCodePath)
    
    'Debug.Print "RunAsExe complete.", Now
    RunAsExe = (m_ExitCodeUpdate = 0)
    '---
    Exit Function
RoutinError:
    Call MessageError(Err)
End Function

'**************
'* PUBLIC SUB *
'**************
Public Sub UnzipFile(pathZip As String, pathDest As String, Optional ByVal deleteZip As Boolean, Optional ByRef fileExtensionUnzip As String, Optional ByRef lastPathUnzipExeFile As String)
    Dim objShell As Shell32.Shell 'Object
    Dim folderZip As Shell32.Folder
    Dim folderDest As Shell32.Folder
    Dim file As Shell32.FolderItem
    Dim nameFile As String
    Dim fileExtension As String

    'Crea una instancia del objeto Shell
    Set objShell = New Shell32.Shell
    
    'Obtener una referencia a la carpeta ZIP
    Set folderZip = objShell.NameSpace(pathZip)
    
    'Obtener una referencia a la carpeta destino
    Set folderDest = objShell.NameSpace(pathDest)
    
    'Validar de que ambas carpetas existen
    If Not folderZip Is Nothing And Not folderDest Is Nothing Then
        'Validar si van a descomprimir por alguna extension especifica.
        If Not fileExtensionUnzip = vbNullString Then
            'Iterando sobre los elementos en la carpeta ZIP
            For Each file In folderZip.Items
                nameFile = GetFileNameFromPath(file.Path)
                'Si el elemento es una carpeta, llamar recursivamente al procedimiento.
                If (file.IsFolder) Then
                    Call UnzipFile(pathZip & "\" & nameFile, pathDest, False, "exe", lastPathUnzipExeFile)
                Else
                    'Obtener la extensión del archivo
                    fileExtension = LCase$(Mid$(nameFile, InStrRev(nameFile, ".") + 1))
                    
                    'Verifica si la extensión del archivo coincide con la extensión específica
                    If fileExtension = LCase$(fileExtensionUnzip) Then
                        'Copia el archivo a la carpeta destino
                        folderDest.CopyHere file, 4
                        'El parámetro 4 es para evitar los mensajes de confirmación (4 = FOF_NO_UI)
                        'Setear variable para devolver el ultimo archivo descomprimido con la extesion especifica.
                        lastPathUnzipExeFile = pathDest & "\" & nameFile
                    End If
                End If
            Next file
        Else
            'Extraer todo el zip.
            folderDest.CopyHere folderZip.Items, 4
        End If
    Else
        MsgBox "Failed to access the ZIP file path or destination folder.", vbExclamation + vbOKOnly
    End If
    
    'Elimina el archivo ZIP después de descomprimir
    If deleteZip Then
        On Error Resume Next
        Kill pathZip
        On Error GoTo 0
    End If
    
    ' Libera los objetos
    Set file = Nothing
    Set folderZip = Nothing
    Set folderDest = Nothing
    Set objShell = Nothing
End Sub

Public Sub MoveFileToPath(ByVal filePath As String, ByVal folderDest As String, Optional ByVal newNameFile As String, Optional ByVal replaceFileDest As Boolean)
    Dim result As Long
    Dim fileDestPath As String
    
    'Validar si se paso un nuevo nombre para el archivo en el destino
    If newNameFile = vbNullString Then
        'Combina la ruta de destino con el nombre original
        fileDestPath = folderDest & "\" & GetFileNameFromPath(filePath)
    Else
        'Combina la ruta de destino con el nuevo nombre
        fileDestPath = folderDest & "\" & newNameFile
    End If
    
    'Validar si el usuario desea reemeplazar el archiv en el destino, esto eliminara el archivo que existe con el mismo nombre.
    If replaceFileDest Then
        ' Elimina el archivo de destino si ya existe
        If FileExists(fileDestPath) Then
            On Error Resume Next
            Kill fileDestPath
            On Error GoTo 0
        End If
    End If
    
    ' Llama a la función de la API MoveFile para mover el archivo con cambio de nombre
    result = MoveFile(filePath, fileDestPath)
    
    If result = 0 Then
        MsgBox "Failed to move file: " & Err.LastDllError, vbExclamation
    End If
End Sub

Public Sub Wait(milliseconds As Long)
    Dim startTime As Single
    Dim elapsed As Single
    On Error GoTo RoutinError
    '---
    startTime = Timer

    Do While elapsed < milliseconds
        DoEvents
        elapsed = (Timer - startTime) * 1000 ' Convertir a milisegundos
    Loop
    DoEvents
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub

Public Sub LoadDataEmojis()
    'https://es.wiktionary.org/wiki/Ap%C3%A9ndice:Caracteres_Unicode/Emojis
    'https://unicode.org/emoji/charts/full-emoji-list.html
    'rostro sonriente (grinning face)
    arrDataEmojis(0, 0) = "[emoji_grinning_face]"
    arrDataEmojis(0, 1) = "[emoji_rostro_sonriente]"
    arrDataEmojis(0, 2) = "U+1F600"
    'cara sonriente con ojos sonrientes (grinning face with smiling eyes)
    arrDataEmojis(1, 0) = "[emoji_grinning_face_with_smiling_eyes]"
    arrDataEmojis(1, 1) = "[emoji_cara_sonriente_con_ojos_sonrientes]"
    arrDataEmojis(1, 2) = "U+1F601"
    'cara con lágrimas de alegría (face with tears of joy)
    arrDataEmojis(2, 0) = "[emoji_face_with_tears_of_joy]"
    arrDataEmojis(2, 1) = "[emoji_cara_con_lagrimas_de_alegria]"
    arrDataEmojis(2, 2) = "U+1F602"
    'cara sonriente con boca abierta (smiling face with open mouth)
    arrDataEmojis(3, 0) = "[emoji_smiling_face_with_open_mouth]"
    arrDataEmojis(3, 1) = "[emoji_cara_sonriente_con_boca_abierta]"
    arrDataEmojis(3, 2) = "U+1F603"
    'cara sonriente con gafas de sol (smiling face with sunglasses)
    arrDataEmojis(4, 0) = "[emoji_smiling_face_with_sunglasses]"
    arrDataEmojis(4, 1) = "[emoji_cara_sonriente_con_gafas_de_sol]"
    arrDataEmojis(4, 2) = "U+1F60E"
End Sub

Public Sub MessageError(objError As ErrObject)
    MsgBox "The following error was encountered in the process." & vbCrLf & vbCrLf & _
           "Error number:   " & objError.Number & vbCrLf & _
           "Description:    " & objError.Description & vbCrLf & _
           "Source:         " & objError.Source
End Sub
