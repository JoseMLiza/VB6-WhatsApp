Attribute VB_Name = "mdlSelenium"
'=========================================================================
'
' Project   : VB6-WhatsApp
' Module    : mdlSelenium.bas
' Author    : Jose Liza (https://github.com/JoseMLiza)
' Based     : SeleniumBasic
'
'=========================================================================

Option Explicit

Public Const SEEdgeUserDataPath    As String = "\bin\SeleniumBasic\data\edge\userdata"
Public Const SEChromeUserDataPath  As String = "\bin\SeleniumBasic\data\chrome\userdata"

'**********************
'* FUNCIONES PRIVADAS *
'**********************
Private Function SE_jsSimulateMouseEvents() As String
    On Error GoTo FunctionError
    Dim strScript As String
    '---
    strScript = "function simulateMouseEvents(element, eventName) {" & vbCrLf
    strScript = strScript & "  const mouseEvent = new MouseEvent(eventName, {" & vbCrLf
    strScript = strScript & "    bubbles: true," & vbCrLf
    strScript = strScript & "    cancelable: true," & vbCrLf
    strScript = strScript & "    view: window" & vbCrLf
    strScript = strScript & "  });" & vbCrLf
    strScript = strScript & "  element.dispatchEvent(mouseEvent);" & vbCrLf
    strScript = strScript & "}" & vbCrLf
    '---
    SE_jsSimulateMouseEvents = strScript
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Private Function SE_jsSetTextInputElement(ByVal strElement As String, ByVal strText As String) As String
    On Error GoTo FunctionError
    Dim strScript As String
    '---
    strScript = "function setTextElement(text) {" & vbCrLf
    strScript = strScript & "   const dataText = new DataTransfer();" & vbCrLf
    strScript = strScript & "   dataText.setData('text', text);" & vbCrLf
    strScript = strScript & "   setTimeout(()=>{},500);" & vbCrLf
    strScript = strScript & "   //Paste text" & vbCrLf
    strScript = strScript & "   setTimeout(() => {" & vbCrLf
    strScript = strScript & "       const event = new ClipboardEvent('paste', {" & vbCrLf
    strScript = strScript & "           clipboardData: dataText," & vbCrLf
    strScript = strScript & "           bubbles: true" & vbCrLf
    strScript = strScript & "       });" & vbCrLf
    strScript = strScript & "       let inputElement = document.querySelector('" & strElement & "')" & vbCrLf
    strScript = strScript & "       simulateMouseEvents(inputElement, 'click');" & vbCrLf
    strScript = strScript & "       inputElement.focus()" & vbCrLf
    strScript = strScript & "       // select old text and replace it with new" & vbCrLf
    strScript = strScript & "       document.execCommand('selectall');" & vbCrLf
    strScript = strScript & "       inputElement.dispatchEvent(event)" & vbCrLf
    strScript = strScript & "   }, 500);" & vbCrLf
    strScript = strScript & "}" & vbCrLf
    strScript = strScript & "setTextElement('" & strText & "');" & vbCrLf
    '---
    SE_jsSetTextInputElement = strScript
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

'**********************
'* FUNCIONES PUBLICAS *
'**********************
Public Function ValidateSelenium(objWebDriver As Selenium.WebDriver) As Boolean
    On Error GoTo FunctionError
    '---
    If objWebDriver Is Nothing Then Err.Raise 5000, "Selenium.WebDriver", m_ArrayMessage(3)
    '---
    ValidateSelenium = True
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function


'***********************
'* SCRIPTS (JAVASCRIPT)*
'***********************
Public Function SE_PageIsReady() As Boolean
    On Error GoTo FunctionError
    Dim sResult As String
    '---
    If objSEWebDriver Is Nothing Then Exit Function
    '---
    Do While sResult <> "complete"
        sResult = objSEWebDriver.ExecuteScript("return document.readyState")
        DoEvents
    Loop
    '---
    SE_PageIsReady = True
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function SE_GetElementExists(ByVal strElement As String) As Boolean
    On Error GoTo FunctionError
    Dim script As String
    '---
    If objSEWebDriver Is Nothing Then Exit Function
    '---
    On Error Resume Next
    If Left(strElement, 2) = "//" Then
        SE_GetElementExists = objSEWebDriver.IsElementPresent(objSEBy.XPath(strElement))
    Else
        SE_GetElementExists = objSEWebDriver.IsElementPresent(objSEBy.Css(strElement))
    End If
    On Error GoTo 0
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function SE_GetAttribute(ByVal strElement As String, ByVal strAttributeName As String) As String
    On Error GoTo FunctionError
    Dim element As WebElement
    '---
    If objSEWebDriver Is Nothing Then Exit Function
    If SE_GetElementExists(strElement) Then
        On Error Resume Next
        If Left(strElement, 2) = "//" Then
            Set element = objSEWebDriver.FindElement(objSEBy.XPath(strElement))
        Else
            Set element = objSEWebDriver.FindElement(objSEBy.Css(strElement))
        End If
        On Error GoTo 0
        If Not element Is Nothing Then
            SE_GetAttribute = element.Attribute(strAttributeName)
        End If
    End If
    '---
    Exit Function
FunctionError:
    Call MessageError(Err)
End Function

Public Function SE_GetTextElement(ByVal strElement As String) As String
    On Error GoTo RoutinError
    Dim element As WebElement
    '---
    If SE_GetElementExists(strElement) Then
        On Error Resume Next
        If Left(strElement, 2) = "//" Then
            Set element = objSEWebDriver.FindElement(objSEBy.XPath(strElement))
        Else
            Set element = objSEWebDriver.FindElement(objSEBy.Css(strElement))
        End If
        On Error GoTo 0
        If Not element Is Nothing Then
            SE_GetTextElement = element.Text
        End If
    End If
    '---
    Exit Function
RoutinError:
    Call MessageError(Err)
End Function

'**************
'* PUBLIC SUB *
'**************
Public Sub SE_ElementClickEvent(ByVal strElement As String, Optional ByVal ClickAndHold As Boolean)
    On Error GoTo RoutinError
    Dim element As WebElement
    '---
    If objSEWebDriver Is Nothing Then Exit Sub
    If SE_GetElementExists(strElement) Then
        On Error Resume Next
        If Left(strElement, 2) = "//" Then
            Set element = objSEWebDriver.FindElement(objSEBy.XPath(strElement))
        Else
            Set element = objSEWebDriver.FindElement(objSEBy.Css(strElement))
        End If
        On Error GoTo 0
        If Not element Is Nothing Then
            If ClickAndHold Then
                element.ClickAndHold
            Else
                element.Click
            End If
        End If
    End If
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub

Public Sub SE_SetTextInputElement(ByVal strElement As String, ByVal strText As String)
    On Error GoTo RoutinError
    Dim element As WebElement
    '---
    If SE_GetElementExists(strElement) Then
        On Error Resume Next
        If Left(strElement, 2) = "//" Then
            Set element = objSEWebDriver.FindElement(objSEBy.XPath(strElement))
        Else
            Set element = objSEWebDriver.FindElement(objSEBy.Css(strElement))
        End If
        On Error GoTo 0
        If Not element Is Nothing Then
            element.SendKeys strText
        End If
    End If
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub

Public Sub SE_ScriptSetTextInputElement(ByVal strElement As String, ByVal strText As String)
    On Error GoTo RoutinError
    Dim strScript As String
    '---
    If objSEWebDriver Is Nothing Then Exit Sub
    If SE_GetElementExists(strElement) Then
        On Error Resume Next
        '---
        strScript = SE_jsSimulateMouseEvents
        strScript = strScript & SE_jsSetTextInputElement(strElement, strText)
        Call objSEWebDriver.ExecuteScript(strScript)
        '---
        On Error GoTo 0
    End If
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub


Public Sub SE_SendKeyElement(ByVal strElement As String, ByVal strKey As String)
    On Error GoTo RoutinError
    Dim element As WebElement
    '---
    If SE_GetElementExists(strElement) Then
        On Error Resume Next
        If Left(strElement, 2) = "//" Then
            Set element = objSEWebDriver.FindElement(objSEBy.XPath(strElement))
        Else
            Set element = objSEWebDriver.FindElement(objSEBy.Css(strElement))
        End If
        On Error GoTo 0
        If Not element Is Nothing Then
            element.SendKeys strKey
        End If
    End If
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub

Public Sub SE_WebBrowserSendKey(ByVal strKey As String)
    On Error GoTo RoutinError
    '---
    If objSEWebDriver Is Nothing Then Exit Sub
    objSEWebDriver.SendKeys strKey
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub
