Attribute VB_Name = "mdlHTMLFunctions"
'=========================================================================
'
' Project   : VB6-WhatsApp
' Module    : mdlHTMLFunctions.bas
' Author    : Jose Liza (https://github.com/JoseMLiza)
'
'=========================================================================

Option Explicit

Public Function IndentHtml(html As String) As String
    'Función para formatear el HTML con indentación usando tabulaciones
    Dim i As Long
    Dim depth As Long
    Dim indent As String
    Dim result As String
    
    depth = 0
    indent = vbTab ' Usar tabulaciones para la indentación
    
    For i = 1 To Len(html)
        If Mid(html, i, 1) = "<" Then
            If Mid(html, i + 1, 1) = "/" Then
                'Disminuir la profundidad
                depth = depth - 1
            End If
            result = result & vbCrLf & String(depth, indent) & Mid(html, i, 1)
            If Mid(html, i + 1, 1) <> "/" Then
                'Aumentar la profundidad
                depth = depth + 1
            End If
        Else
            result = result & Mid(html, i, 1)
        End If
    Next i
    IndentHtml = result
End Function

