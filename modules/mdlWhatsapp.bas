Attribute VB_Name = "mdlWhatsapp"
'=========================================================================
'
' Project   : VB6-WhatsApp
' Module    : mdlWhatsapp.bas
' Author    : Jose Liza (https://github.com/JoseMLiza)
' Based     : WebView2-Binding-(Edge-Chromium) / SeleniumBasic
'
'=========================================================================

Option Explicit

'***********************
'* PUBLIC DECLARATIONS *
'***********************
Public Enum ENM_LANGUAGE
    LNG_ENGLISH
    LNG_SPANISH
End Enum

Public Enum ENM_STATEWHATSAPP
    STATE_PAGE_LOAD         ' 0 - Cuando se carga la url.
    STATE_PAGE_READY        ' 1 - Cuando cargo la pagina inicial.
    STATE_PAGE_VALIDATE     ' 2 - Cuando valida si se debe o no escanear el QR.
    STATE_WAITQR            ' 3 - Si se debe escanear, se carga el contenido del QR.
    STATE_WAITSCANQR        ' 4 - Se cargó el QR, entonces se debe escanear el QR.
    STATE_SCANQR_CANCEL     ' 5 - Se cancela el escaneo del QR.
    STATE_CHATS_LOADING     ' 6 - Se cargan la listas conversaciones.
    STATE_CHATS_READY       ' 7 - Se completo la carga de las conversaciones.
    STATE_CHATS_FIND        ' 8 - Buscar una conversacion.
    STATE_CHATS_FOUND       ' 9 - Conversacion encontrada.
    STATE_CHAT_NEW          '10 - Si la conversacion no existe, se crea una nueva.
    STATE_CONTACT_NOTEXISTS '11 - Contacto no existe.
    STATE_LOGOUT            '12 - Cerrar sesión.
    STATE_CLOSE             '13 - Cerrar aplicación.
End Enum

Public StateWhatsapp                    As ENM_STATEWHATSAPP

Public m_ArrayStatus(13)                As Variant
Public m_ArrayMessage()                 As Variant
Public m_PageIsReady                    As Boolean

Public Const CS_URLWHATSAPPWEB          As String = "https://web.whatsapp.com/"

'************************
'* PRIVATE DECLARATIONS *
'************************
Private Const CS_ID_DIVAPP              As String = "#app"
Private Const CS_QRYSEL_QRCANVAS        As String = "canvas[aria-label*=""QR""]"
'XPath selector
'Private Const CS_ELEMENT_DIVQR          As String = "//*[@id=""app""]/div/div[2]/div[2]/div[2]/div/div[2]/div[2]/div[1]/div[contains(@class, ""_akau"") or contains(@class, ""_19vUU"")]"
'Private Const CS_ELEMENT_DIVQR2         As String = "//*[@id=""app""]/div/div[2]/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div/div[contains(@class, ""_akau"") or contains(@class, ""_19vUU"")]"
'Private Const CS_XPATH_LOADCHAT         As String = "//*[@id=""app""]/div/div[2]/div[3]/progress[contains(@class,""_ak0k"") or contains(@class, ""ZJWuG"")]"
Private Const CS_XPATH_MENU             As String = "//*[@id=""app""]/div/div[3]/div/div[3]/header/header/div/span/div/div[2]/button/span"
Private Const CS_XPATH_MENU2            As String = "//*[@id=""app""]/div/div[2]/div[3]/header/header/div/span/div/span/div[2]/div/span"
Private Const CS_XPATH_CONTACTNOEXITST  As String = "//*[@id=""app""]/div/div[3]/div/div[2]/div[1]/span/div/span/div/div[2]/div[2]/div/span"
Private Const CS_XPATH_ATTACHDOCUMENT   As String = "//*[@id=""app""]/div/span[5]/div/ul/div/div/div[1]/li/div/span" '"//*[@id=""main""]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[1]/li/div/div"
Private Const CS_XPATH_INPUTATTACHDOC   As String = "//input[@accept=""*""]"
Private Const CS_XPATH_ATTACHIMAGE      As String = "//*[@id=""app""]/div/span[5]/div/ul/div/div/div[2]/li/div/span" '"//*[@id=""main""]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[2]/li/div/div"
Private Const CS_XPATH_INPUTATTACHIMG   As String = "//input[@accept=""image/*,video/mp4,video/3gpp,video/quicktime""]"
'Query selector
Private Const CS_QRYSEL_DIVQR           As String = "div[class~=""_akau""], div[class~=""_19vUU""]"
Private Const CS_QRYSEL_LOADCHAT        As String = "progress[class~=""_ak0k""], progress[class~=""ZJWuG""]"
Private Const CS_QRYSEL_DIVPOPUP        As String = "div[data-animate-modal-popup=""true""]" 'Div del mensaje de actualizacion de Whatsapp.
Private Const CS_QRYSEL_DIVPOPUPBTN     As String = "//*[@id=""app""]/div/span[2]/div/div/div/div/div/div/div[2]/div/button" 'Boton del div popup.
Private Const CS_QRYSEL_DIVPOPUPBTN2    As String = "//*[@id=""app""]/div/span[2]/div/div/div/div/div/div/div[2]/div/button/div" 'Boton del div popup.
Private Const CS_QRYSEL_MENU            As String = "[data-icon*=""menu""]"
Private Const CS_QRYSEL_NEWCHAT         As String = "[data-icon*=""new-chat""]"
Private Const CS_QRYSEL_CHATNOEXITST    As String = "#app div#side div#pane-side span._ao3e"
Private Const CS_QRYSEL_NEWCHATRETURN   As String = "[data-icon=""back""]"
Private Const CS_QRYSEL_INPUTNUMBER     As String = "#app [contenteditable=""true""][role=""textbox""]"
Private Const CS_QRYSEL_BTNSEARCH       As String = "[data-icon*=""search""]" 'WhatsApp Personal
Private Const CS_QRYSEL_BTNSEARCH2      As String = "[data-icon*=""search-refreshed""]" 'WhatsApp Business
Private Const CS_QRYSEL_INPUTMESSAGE    As String = "#main .copyable-area [contenteditable=""true""][role=""textbox""]"
Private Const CS_QRYSEL_INPUTCOMMENT    As String = "#app .copyable-area [contenteditable=""true""][role=""textbox""]" 'Para comentarios en los archivos adjuntos.
Private Const CS_QRYSEL_BTNATTACH       As String = "span [data-icon*=""plus""]"
Private Const CS_QRYSEL_BTNADDALT       As String = "span [data-icon*=""add-alt""]"
'Private Const CS_QRYSEL_ATTACHFILE      As String = "#main input[type=""file""][accept=""*""][multiple]"
'Private Const CS_QRYSEL_ATTACHIMAGE     As String = "#main input[type=""file""][accept=""image/*,video/mp4,video/3gpp,video/quicktime""][multiple]"
Private Const CS_QRYSEL_BTNSEND         As String = "#main [data-icon*=""send""]"
'Private Const CS_XPATH_MESSAGETEXT    As String = "//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]"

Private Const CS_IMAGEEXTENSIONS        As String = "xbm|tif|jfif|ico|tiff|gif|svg|jpeg|svgz|jpg|webp|png|bmp|pjp|apng|pjpeg|avif|m4v|mp4|3gp|mov"

Private m_arrImageExtensions()          As String

Private m_WebView2                      As cWebView2
Private m_Language                      As ENM_LANGUAGE

Private m_Files                         As Collection

'*******************
'* PUBLIC PROPERTY *
'*******************
Public Property Get Files() As Collection
    If m_Files Is Nothing Then Set m_Files = New Collection
    Set Files = m_Files
End Property

Public Property Set Files(NewValue As Collection)
    Set m_Files = NewValue
End Property

'**************
'* PUBLIC SUB *
'**************
Public Sub p_Initialize(Optional language As ENM_LANGUAGE = LNG_ENGLISH)
    ReDim m_ArrayMessage(0 To 10)
    m_Language = language
    Select Case m_Language
        Case LNG_ENGLISH
            'ENGLISH
            'Messages
            m_ArrayMessage(0) = "WebView2 object cannot be null"
            m_ArrayMessage(1) = "Login required." & vbCrLf & "Please scan the WhatsApp Web QR code."
            m_ArrayMessage(2) = "QR Updates: "
            m_ArrayMessage(3) = "Selenium.WebDriver object cannot be null"
            m_ArrayMessage(4) = "No contact information specified, please try again."
            m_ArrayMessage(5) = "The message to send has not been specified, try again."
            m_ArrayMessage(6) = "Error searching for contact for new chat."
            'Status
            m_ArrayStatus(0) = "Connecting..."
            m_ArrayStatus(1) = "Connection established."
            m_ArrayStatus(2) = "Validating session..."
            m_ArrayStatus(3) = "Getting QR..."
            m_ArrayStatus(4) = "Scan QR."
            m_ArrayStatus(5) = "QR scan canceled."
            m_ArrayStatus(6) = "Loading conversations..." '-> Valor se actualizar al escanear el QR.
            m_ArrayStatus(7) = "Conversations loaded."
            m_ArrayStatus(8) = "Searching for conversation..."
            m_ArrayStatus(9) = "Conversation found."
            m_ArrayStatus(10) = "Creating new conversation..."
            m_ArrayStatus(11) = "Contact does not exist..."
            m_ArrayStatus(12) = "Closing session..."
            m_ArrayStatus(13) = "Closing the application."
        Case LNG_SPANISH
            'SPANISH
            'Mensajes
            m_ArrayMessage(0) = "El objeto WebView2 no puede ser nulo"
            m_ArrayMessage(1) = "Se requiere inicio de sesión." & vbCrLf & "Por favor escanee el codigo QR de WhatsApp Web."
            m_ArrayMessage(2) = "Actualizaciones del QR: "
            m_ArrayMessage(3) = "El objeto Selenium.WebDriver no puede ser nulo"
            m_ArrayMessage(4) = "No se ha especificado información del contacto, intente nuevamente."
            m_ArrayMessage(5) = "No se ha especificado el mensaje a envíar, intente nuevamente."
            m_ArrayMessage(6) = "Error al buscar contacto para el nuevo chat."
            'Estado
            m_ArrayStatus(0) = "Conectando..."
            m_ArrayStatus(1) = "Conexión establecida."
            m_ArrayStatus(2) = "Validando sesión..."
            m_ArrayStatus(3) = "Obteniendo QR..."
            m_ArrayStatus(4) = "Escanear QR."
            m_ArrayStatus(5) = "Escaneo de QR cancelado."
            m_ArrayStatus(6) = "Cargando conversaciones..." '-> Valor se actualizar al escanear el QR.
            m_ArrayStatus(7) = "Conversaciones cargadas."
            m_ArrayStatus(8) = "Buscando conversación..."
            m_ArrayStatus(9) = "Conversación encontrada."
            m_ArrayStatus(10) = "Creando nueva conversacion..."
            m_ArrayStatus(11) = "Contacto no existe..."
            m_ArrayStatus(12) = "Cerrando sesión..."
            m_ArrayStatus(13) = "Cerrando la aplicación."
    End Select
    '---
    m_arrImageExtensions = Split(CS_IMAGEEXTENSIONS, "|")
    '---
End Sub

Public Sub p_ValidateAccess()
    Dim count As Integer, div As Boolean, login As Boolean
    On Error GoTo RoutinError
    '---
    Select Case webApp
        Case EdgeWebView2
            If Not ValidateWebView2(m_WebView2) Then Exit Sub
            If StateWhatsapp = STATE_PAGE_READY Then
                StateWhatsapp = STATE_PAGE_VALIDATE
                Do
                    count = WebView2_GetChildsCount(m_WebView2, CS_ID_DIVAPP)
                    DoEvents
                Loop While count = 0
                '---
                Do
                    If WebView2_GetElementExists(m_WebView2, CS_QRYSEL_DIVQR) Then
                        div = True: login = False
                    ElseIf WebView2_GetElementExists(m_WebView2, CS_XPATH_MENU) Or WebView2_GetElementExists(m_WebView2, CS_XPATH_MENU2) Then
                        div = True: login = True
                    End If
                    DoEvents
                Loop While div = False
                '---
                If Not login Then
                    StateWhatsapp = STATE_WAITQR
                    frmQRScan.Show vbModal
                Else
                    'Do
                    '    div = WebView2_GetElementExists(m_WebView2, CS_XPATH_MENU)
                    '    DoEvents
                    'Loop While div = False
                    Call p_ValidateChatsReady
                End If
            End If
        Case SeleniumBasic
            If Not ValidateSelenium(objSEWebDriver) Then Exit Sub
            If StateWhatsapp = STATE_PAGE_READY Then
                StateWhatsapp = STATE_PAGE_VALIDATE
                Wait 1000
                '--
                With objSEWebDriver
                    Do
                        If SE_GetElementExists(CS_QRYSEL_DIVQR) Then
                            div = True: login = False
                        ElseIf SE_GetElementExists(CS_XPATH_MENU) Or SE_GetElementExists(CS_XPATH_MENU2) Then
                            div = True: login = True
                        End If
                        DoEvents
                    Loop While div = False
                    If Not login Then
                        StateWhatsapp = STATE_WAITQR
                        frmQRScan.Show vbModal
                    Else
                        'Validar chat ready
                        Call p_ValidateChatsReady
                    End If
                End With
            End If
    End Select
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub

Public Sub SendWhatsapp(ByVal numberContact As String, ByVal message As String)
    On Error GoTo RoutinError
    Dim i As Integer
    Dim strTemp As String
    '---
    If Len(Trim(numberContact)) = 0 Then
        MsgBox m_ArrayMessage(4), vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    If Len(Trim(message)) = 0 Then
        MsgBox m_ArrayMessage(5), vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    '---
    Select Case webApp
        Case EdgeWebView2
            With m_WebView2
                Call .jsRun("newChat", numberContact)
                Wait 1000
                '--
                If WebView2_GetElementExists(m_WebView2, CS_XPATH_CONTACTNOEXITST) Then
                    strTemp = WebView2_GetTextElement(m_WebView2, CS_XPATH_CONTACTNOEXITST)
                    MsgBox m_ArrayMessage(6) & vbCrLf & vbCrLf & strTemp, vbCritical + vbOKOnly
                    Call .jsRun("cancelNewChat")
                Else
                    Call .jsRun("confirmNewChat")
                    Call .jsRun("setTextMessage", message)
                    Call .jsRun("sendButtonMessage")
                    'Confirmar envío - comentado para resolver el problema de ejecucion asincrona
                    'MsgBox "Message was sent successfully.", vbInformation + vbOKOnly
                End If
            End With
        Case SeleniumBasic
            Call SE_ElementClickEvent(CS_QRYSEL_NEWCHAT)
            Wait 1000
            Call SE_ElementClickEvent(CS_QRYSEL_INPUTNUMBER)
            Call SE_SetTextInputElement(CS_QRYSEL_INPUTNUMBER, numberContact)
            Wait 800
            If Not SE_GetElementExists(CS_XPATH_CONTACTNOEXITST) Then
                Call SE_SendKeyElement(CS_QRYSEL_INPUTNUMBER, objKeys.Enter)
            Else
                MsgBox m_ArrayMessage(6) & vbCrLf & vbCrLf & _
                       SE_GetTextElement(CS_XPATH_CONTACTNOEXITST), vbCritical + vbOKOnly
                Wait 300
                Call SE_ElementClickEvent(CS_QRYSEL_NEWCHATRETURN)
                Exit Sub
            End If
            If Not m_Files Is Nothing Then
                If m_Files.count > 0 Then
                    Call SE_ElementClickEvent(CS_QRYSEL_BTNATTACH)
                    Wait 500 'Esperar que abra la lista de tipo de adjuntos
                    If f_ValidateImageExtension(m_Files.Item(1)) Then
                        'IMAGENES
                        Call SE_ElementClickEvent(CS_XPATH_ATTACHIMAGE, True)
                        Wait 500
                        Call SE_SendKeyElement(CS_XPATH_INPUTATTACHIMG, m_Files.Item(1))
                    Else
                        'DOCUMENTOS
                        Call SE_ElementClickEvent(CS_XPATH_ATTACHDOCUMENT, True)
                        Wait 500
                        Call SE_SendKeyElement(CS_XPATH_INPUTATTACHDOC, m_Files.Item(1))
                    End If
                    Wait 2000 'Esperar cargar del archivo
                    If m_Files.count > 1 Then
                        For i = 2 To m_Files.count 'Agregar archivos restantes
                            Call SE_ElementClickEvent(CS_QRYSEL_BTNADDALT, True)
                            Wait 500
                            Call SE_SendKeyElement(CS_XPATH_INPUTATTACHDOC, m_Files.Item(i))
                            Wait 2000 'Esperar cargar del archivo
                        Next
                    End If
                    If Len(Trim(message)) > 0 Then
                        Call SE_ElementClickEvent(CS_QRYSEL_INPUTCOMMENT)
                        'Call SE_SetTextInputElement(CS_QRYSEL_INPUTCOMMENT, message)
                        Call SE_ScriptSetTextInputElement(CS_QRYSEL_INPUTCOMMENT, message)
                        Wait 500
                        'Enviar mensaje
                        Call SE_SendKeyElement(CS_QRYSEL_INPUTCOMMENT, objKeys.Enter)
                        'Confirmar envío.
                        MsgBox "Message was sent successfully.", vbInformation + vbOKOnly
                    End If
                End If
            Else
                Wait 500
                Call SE_ElementClickEvent(CS_QRYSEL_INPUTMESSAGE)
                'Call SE_SetTextInputElement(CS_QRYSEL_INPUTMESSAGE, message)
                Call SE_ScriptSetTextInputElement(CS_QRYSEL_INPUTMESSAGE, message)
                Wait 500
                'Enviar mensaje
                Call SE_SendKeyElement(CS_QRYSEL_INPUTMESSAGE, objKeys.Enter)
                'Confirmar envío.
                MsgBox "Message was sent successfully.", vbInformation + vbOKOnly
            End If
    End Select
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub

Public Sub p_WebView2OpenWhatsappWeb(ByVal wvBrowser As cWebView2)
    Set m_WebView2 = wvBrowser
    If ValidateWebView2(m_WebView2) Then
        Call AttachScripts(m_WebView2)
        With m_WebView2
            'Agregar el script para crear un nuevo chat.
            .AddScriptToExecuteOnDocumentCreated ScriptNewChat 'Script JS: newChat
            .AddScriptToExecuteOnDocumentCreated ScriptValidateNewChat 'Script JS: validateNewChat
            .AddScriptToExecuteOnDocumentCreated ScriptConfirmNewChat 'Script JS: confirmNewChat
            .AddScriptToExecuteOnDocumentCreated ScriptCancelNewChat 'Script JS: cancelNewChat
            
            'Agregar el script para setear texto en un nuevo mensaje.
            .AddScriptToExecuteOnDocumentCreated ScriptSetTextMessage 'Script JS: setTextMessage
            
            'Agregar el script para adjuntar archivo al mensaje.
            '.AddScriptToExecuteOnDocumentCreated ScriptAttachFile 'Script JS: attachFile
            
            'Agregar el script para realizar el envio del mensaje.
            .AddScriptToExecuteOnDocumentCreated ScriptSendMessage 'Script JS: sendButtonMessage
            
            'Abrir la url de WhatsApp Web
            .Navigate CS_URLWHATSAPPWEB
        End With
        StateWhatsapp = STATE_PAGE_LOAD
    End If
End Sub

'Public Sub p_SEOpenWhatsappWeb(ByVal wbDriver As Object)
Public Sub p_SEOpenWhatsappWeb(ByVal wbDriver As Selenium.WebDriver)
    Dim argsWebBrowser As String, argUserAgent As String, strBinaryPath As String
    '---
    strBinaryPath = App.Path & SeleniumBase & IIf(SEWebDriver = EdgeDriver, "edgedriver.exe", "chromedriver.exe")
    argUserAgent = Replace(IIf(SEWebDriver = EdgeDriver, UserAgentMsEdge, UserAgentChrome), "{VERSION}", IIf(m_wdVersion = vbNullString, m_wbVersion, m_wdVersion))
    m_UserDataDir = Replace(App.Path & IIf(SEWebDriver = EdgeDriver, SEEdgeUserDataPath, SEChromeUserDataPath), "\", "\\")
    If m_IsHeadless Then
        argsWebBrowser = "{" & """args"":[""user-data-dir=" & m_UserDataDir & """, ""profile-directory=VB6"", ""no-first-run"", ""disable-infobars"", ""disable-session-crashed-bubble"", ""headless"", ""user-agent=" & argUserAgent & """]}"
    Else
        argsWebBrowser = "{" & """args"":[""user-data-dir=" & m_UserDataDir & """, ""profile-directory=VB6"", ""no-first-run"", ""disable-infobars"", ""disable-session-crashed-bubble"", ""user-agent=" & argUserAgent & """]}"
    End If
    '---
    If Not wbDriver Is Nothing Then
        With wbDriver
            .SetBinaryWebDriver strBinaryPath
            'Setear carpeta de perfil de usuario y abrir en navegador sin cabeza (oculto)
            '.SetCapability "ms:edgeOptions", argsWebBrowser
            .SetCapability IIf(SEWebDriver = EdgeDriver, "ms:edgeOptions", "goog:chromeOptions"), argsWebBrowser
            'Inicra el navegador
            StateWhatsapp = STATE_PAGE_LOAD
            Wait 1000
            .Start
            'Abrir pagina web
            .Get CS_URLWHATSAPPWEB
        End With
        '---
        m_PageIsReady = SE_PageIsReady
        If m_PageIsReady Then
            StateWhatsapp = STATE_PAGE_READY
            Call p_ValidateAccess
        End If
    End If
End Sub

Public Sub p_ValidateChatsReady()
    Dim div As Boolean
    Dim element As WebElement
    On Error GoTo RoutinError
    '---
    Select Case webApp
        Case EdgeWebView2
            'Loop hasta que se cargue el div del menu para saber que se completo la carga de las conversaciones.
            Do
                div = WebView2_GetElementExists(m_WebView2, CS_XPATH_MENU) Or WebView2_GetElementExists(m_WebView2, CS_XPATH_MENU2)
                DoEvents
            Loop While div = False
            StateWhatsapp = STATE_CHATS_READY
            'Loop para validar si se muestra mensajes popup de WhatsApp
            div = False
            Do
                div = WebView2_GetElementExists(m_WebView2, CS_QRYSEL_DIVPOPUP) And WebView2_GetElementExists(m_WebView2, CS_QRYSEL_DIVPOPUPBTN)
                DoEvents
            Loop While div = False
            If div Then Call WebView2_SimulateMouseEvents(m_WebView2, CS_QRYSEL_DIVPOPUPBTN, "click")
        Case SeleniumBasic
            'Loop hasta que se cargue el div del menu para saber que se completo la carga de las conversaciones.
            Do
                div = SE_GetElementExists(CS_XPATH_MENU) Or SE_GetElementExists(CS_XPATH_MENU2)
                DoEvents
            Loop While div = False
            StateWhatsapp = STATE_CHATS_READY
            'Loop para validar si se muestra mensajes popup de WhatsApp
            div = False
            Do
                div = SE_GetElementExists(CS_QRYSEL_DIVPOPUP) And SE_GetElementExists(CS_QRYSEL_DIVPOPUPBTN)
                DoEvents
            Loop While div = False
            If div Then Call SE_ElementClickEvent(CS_QRYSEL_DIVPOPUPBTN)
    End Select
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub

'********************
'* PRIVATE FUNCTION *
'********************
Private Function f_ValidateImageExtension(ByVal pathFile As String) As Boolean
    On Error GoTo FunctionError
    Dim i As Integer
    Dim fileExtension As String
    '---
    fileExtension = LCase(Mid(pathFile, InStrRev(pathFile, ".") + 1))
    '---
    f_ValidateImageExtension = False
    For i = 0 To UBound(m_arrImageExtensions)
        If fileExtension = m_arrImageExtensions(i) Then
            f_ValidateImageExtension = True
            Exit For
        End If
    Next
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

'*******************
'* PUBLIC FUNCTION *
'*******************
'>>>>>>>>>>>>
'>> SCRIPT >>
'>>>>>>>>>>>>
Public Function f_GetDataQR(Optional IsValidate As Boolean = False) As String
    Dim strDataQR As String
    On Error GoTo FunctionError
    '---
    Select Case webApp
        Case EdgeWebView2
            If Not ValidateWebView2(m_WebView2) Then Exit Function
            If WebView2_GetElementExists(m_WebView2, CS_QRYSEL_DIVQR) Then
                If Not IsValidate Then StateWhatsapp = STATE_WAITQR 'Estado: Esperando la carga del QR
                Do
                    strDataQR = WebView2_GetAttributeElement(m_WebView2, CS_QRYSEL_DIVQR, "data-ref")
                    DoEvents
                Loop While Len(strDataQR) = 0
                f_GetDataQR = strDataQR
                If Not IsValidate Then StateWhatsapp = STATE_WAITSCANQR 'Estado: Cargo el QR, pero esperando que se escanee.
            End If
        Case SeleniumBasic
            If Not ValidateSelenium(objSEWebDriver) Then Exit Function
            If SE_GetElementExists(CS_QRYSEL_DIVQR) Then
                If Not IsValidate Then StateWhatsapp = STATE_WAITQR 'Estado: Esperando la carga del QR
                Do
                    strDataQR = SE_GetAttribute(CS_QRYSEL_DIVQR, "data-ref")
                    DoEvents
                Loop While Len(strDataQR) = 0
                f_GetDataQR = strDataQR
                If Not IsValidate Then StateWhatsapp = STATE_WAITSCANQR 'Estado: Cargo el QR, pero esperando que se escanee.
            End If
    End Select
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function f_ValidateScanQR() As Boolean
    Dim div As Boolean
    On Error GoTo FunctionError
    '---
    Select Case webApp
        Case EdgeWebView2
            If Not ValidateWebView2(m_WebView2) Then Exit Function
            'Debug.Print "Escaneo de QR: " & Not (f_ExistsElement(CS_CLASS_DIVQR))
            f_ValidateScanQR = Not WebView2_GetElementExists(m_WebView2, CS_QRYSEL_DIVQR)
            If f_ValidateScanQR Then
                'Loop hasta que se cargue el div que muestra la carga de las conversaciones.
                Do
                    div = WebView2_GetElementExists(m_WebView2, CS_QRYSEL_LOADCHAT)
                    DoEvents
                Loop While div = False
                '---
                If div Then
                    'm_ArrayStatus(6) = f_WebViewEval("document.querySelector('" & CS_CLASS_DIVQR & "').innerText")
                    StateWhatsapp = STATE_CHATS_LOADING
                End If
            End If
        Case SeleniumBasic
            If Not ValidateSelenium(objSEWebDriver) Then Exit Function
            f_ValidateScanQR = Not SE_GetElementExists(CS_QRYSEL_DIVQR)
            If f_ValidateScanQR Then
                'Loop hasta que se cargue el div que muestra la carga de las conversaciones.
                Do
                    div = SE_GetElementExists(CS_QRYSEL_LOADCHAT)
                    DoEvents
                Loop While div = False
                '---
                If div Then
                    StateWhatsapp = STATE_CHATS_LOADING
                End If
            End If
    End Select
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function f_WebViewEval(ByVal Value As String) As String
    f_WebViewEval = m_WebView2.jsRun("eval", Value)
End Function

'*********************
'* PRIVATE FUNCTIONS *
'*********************
Private Function ScriptNewChat() As String
    With New_c.StringBuilder
        .AddNL "function newChat(numberContact, message) {"
        .AddNL "   const dataContact = new DataTransfer();"
        .AddNL "   dataContact.setData('text', numberContact);"
        .AddNL "   "
        .AddNL "   //Open contextmenu NEW CHAT"
        .AddNL "   var newchat = document.querySelector('" & CS_QRYSEL_NEWCHAT & "');"
        .AddNL "   newchat.click();"
        .AddNL "   "
        .AddNL "   //Setfocus text search element"
        .AddNL "   setTimeout(() => {"
        .AddNL "       const event = new ClipboardEvent('paste', {"
        .AddNL "           clipboardData: dataContact,"
        .AddNL "           bubbles: true"
        .AddNL "       });"
        .AddNL "       let el = document.querySelector('" & CS_QRYSEL_INPUTNUMBER & "')"
        .AddNL "       simulateMouseEvents(el, 'mousedown');"
        .AddNL "       el.focus()"
        .AddNL "       "
        .AddNL "       var btnsearch = document.querySelector('" & CS_QRYSEL_BTNSEARCH & "');"
        .AddNL "       if (btnsearch === null){"
        .AddNL "           btnsearch = document.querySelector('" & CS_QRYSEL_BTNSEARCH2 & "');"
        .AddNL "       }"
        .AddNL "       simulateMouseEvents(btnsearch, 'mousedown');"
        .AddNL "       "
        .AddNL "       // select old text and replace it with new"
        .AddNL "       setTimeout(() => {"
        .AddNL "           document.execCommand('selectall');"
        .AddNL "           el.dispatchEvent(event);"
        .AddNL "       },500);"
        .AddNL "   },500);"
        .AddNL "}"
        
        ScriptNewChat = .ToString
    End With
End Function

Private Function ScriptValidateNewChat() As String
    With New_c.StringBuilder
        .AddNL "function validateNewChat(){"
        .AddNL "  //Validate conatct"
        .AddNL "  setTimeout(() => {"
        .AddNL "    if (getElementExits('" & CS_XPATH_CONTACTNOEXITST & "')=== true) {"
        .AddNL "      let textError = getTextElement('" & CS_XPATH_CONTACTNOEXITST & "');"
        .AddNL "      console.log(textError);"
        .AddNL "      //Retornar mensaje de error"
        .AddNL "      return textError;"
        .AddNL "    }else{"
        .AddNL "      return '';"
        .AddNL "    }"
        .AddNL "  }, 1000);"
        .AddNL "}"
        
        ScriptValidateNewChat = .ToString
    End With
End Function

Private Function ScriptCancelNewChat() As String
    With New_c.StringBuilder
        .AddNL "function cancelNewChat(){"
        .AddNL "  var btnBack = document.querySelector('" & CS_QRYSEL_NEWCHATRETURN & "');"
        .AddNL "  simulateMouseEvents(btnBack, 'click');"
        .AddNL "}"
        
        ScriptCancelNewChat = .ToString
    End With
End Function

Private Function ScriptConfirmNewChat() As String
    With New_c.StringBuilder
        .AddNL "function confirmNewChat(){"
        .AddNL "  //Confirm contact"
        .AddNL "  setTimeout(() => {"
        .AddNL "    // Simulate pressing the Enter key"
        .AddNL "    let el = document.querySelector('" & CS_QRYSEL_INPUTNUMBER & "')"
        .AddNL "    const enterKeyEvent = new KeyboardEvent('keydown', {"
        .AddNL "      key: 'Enter',"
        .AddNL "      code: 'Enter',"
        .AddNL "      keyCode: 13,"
        .AddNL "      which: 13,"
        .AddNL "    });"
        .AddNL "    el.dispatchEvent(enterKeyEvent);"
        .AddNL "  }, 500);"
        .AddNL "}"
        
        ScriptConfirmNewChat = .ToString
    End With
End Function

Private Function ScriptSetTextMessage() As String
    With New_c.StringBuilder
        .AddNL "function setTextMessage(message) {"
        .AddNL "   const dataMessage = new DataTransfer();"
        .AddNL "   dataMessage.setData('text', message);"
        '.AddNL "   setTimeout(()=>{},1000);"
        .AddNL "   //Paste message"
        .AddNL "   setTimeout(() => {"
        .AddNL "       const event = new ClipboardEvent('paste', {"
        .AddNL "           clipboardData: dataMessage,"
        .AddNL "           bubbles: true"
        .AddNL "       });"
        .AddNL "       let msgEl = document.querySelector('" & CS_QRYSEL_INPUTMESSAGE & "')"
        .AddNL "       simulateMouseEvents(msgEl, 'click');"
        .AddNL "       msgEl.focus()"
        .AddNL "       // select old text and replace it with new"
        .AddNL "       document.execCommand('selectall');"
        .AddNL "       msgEl.dispatchEvent(event)"
        .AddNL "   }, 1000);"
        .AddNL "}"
        
        ScriptSetTextMessage = .ToString
    End With
End Function

'Private Function ScriptAttachFile() As String
'    With New_c.StringBuilder
'        .AddNL "function attachFile(){"
'        .AddNL "    setTimeout(()=>{},1500);"
'        .AddNL "    setTimeout(()=>{"
'        .AddNL "        document.querySelector('" & CS_QRYSEL_BTNATTACH & "').click();"
'        .AddNL "        setTimeout(()=>{"
'        .AddNL "            //document.querySelector('" & CS_QRYSEL_ATTACHFILE & "').click();"
'        .AddNL "            var addFile = document.querySelector('" & CS_QRYSEL_ATTACHFILE & "');"
'        .AddNL "            simulateMouseEvents(addFile, 'click');"
'        .AddNL "        },1500);"
'        .AddNL "    },3000);"
'        .AddNL "}"
'
'        ScriptAttachFile = .ToString
'    End With
'End Function

Private Function ScriptSendMessage() As String
    With New_c.StringBuilder
        .AddNL "function sendButtonMessage() {"
        .AddNL "    //Send message"
        .AddNL "    setTimeout(()=>{},500);" 'Modificar si se adjuntan archivos
        .AddNL "    setTimeout(() => {"
        .AddNL "        const sendButton ="
        .AddNL "        document.querySelector('" & CS_QRYSEL_BTNSEND & "');"
        .AddNL "    if (sendButton) {"
        .AddNL "        //sendButton.click();"
        .AddNL "        simulateMouseEvents(sendButton, 'click')"
        .AddNL "    } else {"
        .AddNL "        console.error('Send button not found');"
        .AddNL "    }"
        .AddNL "    }, 2500);"
        .AddNL "}"
        ScriptSendMessage = .ToString
    End With
End Function
