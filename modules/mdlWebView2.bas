Attribute VB_Name = "mdlWebView2"
'=========================================================================
'
' Project   : VB6-WhatsApp
' Module    : mdlWebView2.bas
' Author    : Jose Liza (https://github.com/JoseMLiza)
' Based     : WebView2-Binding-(Edge-Chromium)
'
'=========================================================================

Option Explicit

Public Const WebView2UserDataPath    As String = "\bin\RC6\data\userdata"

'********************
'* RUTINAS PUBLICAS *
'********************
Public Sub AttachScripts(objWebView As cWebView2)
    On Error GoTo RoutinError
    '---
    If ValidateWebView2(objWebView) Then
        With objWebView
            'Agregar funcion que permite buscar elementos por tagname, id/classname y xpath.
            .AddScriptToExecuteOnDocumentCreated ScriptFindElement 'Script JS: getElementExits
            
            'Agregar funcion que permite obtener el classname desde el XPath.
            .AddScriptToExecuteOnDocumentCreated ScriptClassFromXPath 'Script JS: getClassByXPath
            
            'Agregar funcion que devulva cuantos elementos hijos tiene en su interior.
            .AddScriptToExecuteOnDocumentCreated ScriptChildsInElement 'Script JS: getChildsInElement
            
            'Agregar funcion que simula eventos del mouse en los elementos.
            .AddScriptToExecuteOnDocumentCreated ScriptSimulateMouseEvents 'Script JS: simulateMouseEvent
            
            'Agregar funcion que permite obtener textos de elementos (input, textarea, div, span y otros).
            .AddScriptToExecuteOnDocumentCreated ScriptGetTextElement 'Scrip JS: getTextElement
            
            'Agregar funcion que permite obtener el atributo de un elemento
            .AddScriptToExecuteOnDocumentCreated ScriptGetAttribute 'Script JS: getAttributeElement
        End With
    End If
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub

'**********************
'* FUNCIONES PUBLICAS *
'**********************
Public Function ValidateWebView2(objWebView As cWebView2) As Boolean
    On Error GoTo FunctionError
    '---
    'Validar si objeto WebView2 es null
    If objWebView Is Nothing Then Err.Raise 5000, "Browser.WebView2", m_ArrayMessage(0) '"WebView2 object cannot be null"
    '---
    ValidateWebView2 = True
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function


Public Function WebView2_GetElementExists(objWebView As cWebView2, ByVal sQuery As String) As Boolean
    On Error GoTo FunctionError
    '---
    WebView2_GetElementExists = objWebView.jsRun("getElementExits", sQuery)
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function WebView2_GetClassFromXPath(objWebView As cWebView2, ByVal sQuery As String) As String
    On Error GoTo FunctionError
    '---
    WebView2_GetClassFromXPath = objWebView.jsRun("getClassByXPath", sQuery)
    If Len(WebView2_GetClassFromXPath) Then
        WebView2_GetClassFromXPath = "." & WebView2_GetClassFromXPath
        WebView2_GetClassFromXPath = Replace(WebView2_GetClassFromXPath, " ", ".")
    End If
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function WebView2_GetChildsCount(objWebView As cWebView2, ByVal sQuery As String) As Integer
    On Error GoTo FunctionError
    '---
    WebView2_GetChildsCount = objWebView.jsRun("getChildsInElement", sQuery)
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function WebView2_GetAttributeElement(objWebView As cWebView2, ByVal element As String, ByVal sQuery As String) As String
    On Error GoTo FunctionError
    '---
    If WebView2_GetElementExists(objWebView, element) Then
        WebView2_GetAttributeElement = objWebView.jsRun("GetAttributeElement", element, sQuery)
    Else
        WebView2_GetAttributeElement = vbNullString
    End If
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function WebView2_GetTextElement(objWebView As cWebView2, ByVal sQuery As String) As String
    On Error GoTo FunctionError
    '---
    WebView2_GetTextElement = objWebView.jsRun("getTextElement", sQuery)
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Sub WebView2_SimulateMouseEvents(objWebView As cWebView2, ByVal element As String, ByVal eventName As String)
    On Error GoTo RoutinError
    '---
    Call objWebView.jsRun("simulateMouseEvents", element, eventName)
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub

'***********************
'* SCRIPTS (JAVASCRIPT)*
'***********************

Private Function ScriptFindElement() As String
    On Error GoTo FunctionError
    '---
    With New_c.StringBuilder
        .AddNL "function getElementExits(sQuery) {"
        .AddNL "    var element = null;"
        .AddNL "    if (sQuery.startsWith('//')) {"
        .AddNL "        var xpathResult = document.evaluate(sQuery, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);"
        .AddNL "        element = xpathResult.singleNodeValue;"
        .AddNL "    } else {"
        .AddNL "        element = document.querySelector(sQuery);"
        .AddNL "    }"
        .AddNL "    return element !== null;"
        .AddNL "}"
        ScriptFindElement = .ToString
    End With
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Private Function ScriptClassFromXPath() As String
    On Error GoTo FunctionError
    '---
    With New_c.StringBuilder
        .AddNL "function getClassByXPath(sQuery) {"
        .AddNL "    var element = document.evaluate(sQuery, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;"
        .AddNL "    if (element) {"
        .AddNL "        return element.className;"
        .AddNL "    } else {"
        .AddNL "        return null;"
        .AddNL "    }"
        .AddNL "}"
        ScriptClassFromXPath = .ToString
    End With
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Private Function ScriptChildsInElement() As String
    On Error GoTo FunctionError
    '---
    With New_c.StringBuilder
        .AddNL "function getChildsInElement(sQuery) {"
        .AddNL "    return document.querySelector(sQuery).childElementCount;"
        .AddNL "}"
        ScriptChildsInElement = .ToString
    End With
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Private Function ScriptSimulateMouseEvents() As String
    On Error GoTo FunctionError
    '---
    With New_c.StringBuilder
'        .AddNL "function simulateMouseEvents(element, eventName) {"
'        .AddNL "  var mouseEvent = document.createEvent('MouseEvents');"
'        .AddNL "  mouseEvent.initEvent(eventName, true, true);"
'        .AddNL "  element.dispatchEvent(mouseEvent);"
'        .AddNL "}"
        .AddNL "function simulateMouseEvents(element, eventName) {"
        .AddNL "  const mouseEvent = new MouseEvent(eventName, {"
        .AddNL "    bubbles: true,"
        .AddNL "    cancelable: true,"
        .AddNL "    view: window"
        .AddNL "  });"
        .AddNL "  element.dispatchEvent(mouseEvent);"
        .AddNL "}"
        ScriptSimulateMouseEvents = .ToString
    End With
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Private Function ScriptGetTextElement() As String
    On Error GoTo FunctionError
    '---
    With New_c.StringBuilder
        .AddNL "function getTextElement(sQuery) {"
        .AddNL "    var element = null;"
        .AddNL "    var textReturn = '';"
        .AddNL "    if (sQuery.startsWith('//')) {"
        .AddNL "        var xpathResult = document.evaluate(sQuery, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);"
        .AddNL "        element = xpathResult.singleNodeValue;"
        .AddNL "    } else {"
        .AddNL "        element = document.querySelector(sQuery);"
        .AddNL "    }"
        .AddNL "    if (element.tagName ==='INPUT' || element.tagName === 'TEXTAREA'){"
        .AddNL "        textReturn = element.value;"
        .AddNL "    }else{"
        .AddNL "        textReturn = element.innerText;"
        .AddNL "    }"
        .AddNL "    return textReturn;"
        .AddNL "}"
        ScriptGetTextElement = .ToString
    End With
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Private Function ScriptGetAttribute() As String
    On Error GoTo FunctionError
    '---
    With New_c.StringBuilder
        .AddNL "function GetAttributeElement(selector, nombreAtributo) {"
        .AddNL "    var elemento;"
        .AddNL "   "
        .AddNL "    // Intentar seleccionar el elemento por xpath"
        .AddNL "    if (selector.startsWith('//') || selector.startsWith('(//')) {"
        .AddNL "        var resultado = document.evaluate(selector, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);"
        .AddNL "        if (resultado && resultado.singleNodeValue) {"
        .AddNL "            elemento = resultado.singleNodeValue;"
        .AddNL "        }"
        .AddNL "    } else {"
        .AddNL "        elemento = document.querySelector(selector);"
        .AddNL "    }"
        .AddNL "   "
        .AddNL "    // Verificar si se encontró un elemento"
        .AddNL "    if (elemento) {"
        .AddNL "        // Verificar si el elemento tiene el atributo especificado"
        .AddNL "        if (elemento.hasAttribute(nombreAtributo)) {"
        .AddNL "            // Devolver el valor del atributo"
        .AddNL "            return elemento.getAttribute(nombreAtributo);"
        .AddNL "        }"
        .AddNL "    }"
        .AddNL "}"
        
        ScriptGetAttribute = .ToString
    End With
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function
