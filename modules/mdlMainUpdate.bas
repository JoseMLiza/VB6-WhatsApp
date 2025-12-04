Attribute VB_Name = "mdlMainUpdate"
'=========================================================================
'
' Project   : WebDriver-Update
' Module    : mdlMainUpdate.bas
' Author    : Jose Liza (https://github.com/JoseMLiza)
'
'=========================================================================

Option Explicit

'***********************
'* PUBLIC DECLARATIONS *
'***********************
Public Const KB As Double = 1024
Public Const MB As Double = KB * KB
Public Const GB As Double = MB * KB

Public Enum enmTypeMessage
    TypeInformation
    TypeCritical
    TypeExclamation
End Enum

Public Enum enmExitCode
    Ok                  'Exito.
    Cancel              'Cancelado por el usuario.
    WrongURL            'URL errada.
    InaccessibleURL     'URL inaccesible.
    DownloadError       'Error en la descarga.
    WriteFileError      'Error al escribir el archivo.
End Enum

Public m_ArgTitle           As String
Public m_ArgDownloadUrl     As String
Public m_ArgDownloadFolder  As String
Public m_ArgDownloadFile    As String
Public m_ArgUnZipPath       As String
Public m_ExitCode           As enmExitCode

'************************
'* PRIVATE DECLARATIONS *
'************************
'Kernel32.dll
Private Declare Function AllocConsole Lib "Kernel32.dll" () As Long
Private Declare Function FreeConsole Lib "Kernel32.dll" () As Long
Private Declare Function WriteConsole Lib "Kernel32.dll" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, ByVal lpReserved As Long) As Long
Private Declare Function GetStdHandle Lib "Kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function AttachConsole Lib "Kernel32.dll" (ByVal dwProcessId As Long) As Long
Private Declare Function GetConsoleWindow Lib "Kernel32.dll" () As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

'User32.dll
Private Declare Function SendMessage Lib "User32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSysColor Lib "User32.dll" (ByVal nIndex As Long) As Long

'UxTheme.dll
Private Declare Function OpenThemeData Lib "UxTheme.dll" (ByVal hwnd As Long, ByVal pszClassList As String) As Long
Private Declare Function CloseThemeData Lib "UxTheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetThemeColor Lib "UxTheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, ByRef pColor As Long) As Long

'Constantes para User32(SendMessage y GetSysColor)
Private Const PBM_SETBARCOLOR   As Long = &H409
Private Const PBM_SETBKCOLOR    As Long = &H2001
Private Const PROGBAR_DEF_COLOR As Long = &HFF000000 '&H8000000D
Private Const COLOR_HIGHLIGHT   As Long = 13
Private Const COLOR_3DFACE      As Long = 15

'Constantes para UxTheme
Private Const PBFS_BAR          As Long = 1
Private Const PBS_NORMAL        As Long = 1
Private Const TMT_FILLCOLOR     As Long = 3802

Private objShell    As Shell 'Object

Private args()      As String
Private hConsole    As Long
Private chrWritten  As Long
Private cmdAttached As Boolean
Private i           As Integer

Private Const STD_OUTPUT_HANDLE As Long = -11&
Private Const ATTACH_PARENT_PROCESS As Long = -1&

Sub Main()
    On Error GoTo RoutinError
    Dim typeError   As enmTypeMessage
    Dim isError     As Boolean
    Dim fileResPath As String
    '---
    If App.LogMode <> 0 Then
        'Intenta adjuntar a la consola del proceso padre (línea de comandos)
        cmdAttached = AttachConsole(ATTACH_PARENT_PROCESS)
        If Not cmdAttached Then
            'Si no se puede adjuntar, crea una nueva consola
            'AllocConsole
            'hConsole = GetStdHandle(STD_OUTPUT_HANDLE)
        Else
            hConsole = GetStdHandle(STD_OUTPUT_HANDLE)
        End If
        
        args = ParseCommandLine(Command$)
        'args = Split(Command$, " ")
        
        'Procesa los argumentos
        For i = 0 To UBound(args)
            Select Case LCase(args(i))
                Case "--title"
                    If i + 1 <= UBound(args) Then
                        m_ArgTitle = RemoveQuotes(args(i + 1))
                        If InStr(1, m_ArgTitle, "--") > 0 Then isError = True
                        i = i + 1 ' Saltar al siguiente argumento
                    Else
                        isError = True
                    End If
                    If isError Then
                        ShowMessage "Error: Missing value for --title argument", hConsole, cmdAttached
                        Exit Sub
                    End If
                Case "--url"
                    If i + 1 <= UBound(args) Then
                        m_ArgDownloadUrl = RemoveQuotes(args(i + 1))
                        If InStr(1, m_ArgDownloadUrl, "--") > 0 Then isError = True
                        i = i + 1 ' Saltar al siguiente argumento
                    Else
                        isError = True
                    End If
                    If isError Then
                        ShowMessage "Error: Missing value for --url argument", hConsole, cmdAttached
                        Exit Sub
                    End If
                Case "--path"
                    If i + 1 <= UBound(args) Then
                        m_ArgDownloadFolder = RemoveQuotes(args(i + 1))
                        If InStr(1, m_ArgDownloadFolder, "--") > 0 Then isError = True
                        i = i + 1 ' Saltar al siguiente argumento
                    Else
                        isError = True
                    End If
                    If isError Then
                        ShowMessage "Error: Missing value for --path argument", hConsole, cmdAttached
                        Exit Sub
                    End If
                Case "/unzip"
                    If i + 1 <= UBound(args) Then
                        m_ArgUnZipPath = RemoveQuotes(args(i + 1))
                        i = i + 1 ' Saltar al siguiente argumento
                    End If
                    If m_ArgUnZipPath = vbNullString Then
                        m_ArgUnZipPath = m_ArgDownloadFolder & "\" & GetFileNameFromURL(m_ArgDownloadUrl, True)
                    End If
            End Select
        Next i
        
        'Verifica que ambos argumentos se hayan proporcionado
        If m_ArgDownloadUrl = vbNullString Or m_ArgDownloadFolder = vbNullString Then
            ShowMessage "Use: update.exe --url ""<URL_DOWNLOAD_FILE>"" --path ""<FOLDER_PATH>""", hConsole, cmdAttached, TypeCritical
            Exit Sub
        End If
    Else 'Modo desarrollo, siempre EdgeDriver
        'm_ArgTitle = "Modo: Desarrollo"
        'm_ArgDownloadUrl = "https://storage.googleapis.com/chrome-for-testing-public/125.0.6422.141/win32/chromedriver-win32.zip"
        'm_ArgDownloadFolder = ""
    End If
    If Not m_ArgTitle = vbNullString Then frmDownload.Caption = m_ArgTitle
    frmDownload.Show
    fileResPath = Environ$("TEMP") & "\" & App.EXEName & "_exitcode.update"
    Call WriteFileExitCode(m_ExitCode, fileResPath)
    Exit Sub
RoutinError:
    Call MessageError(Err, typeError)
End Sub

Public Sub WriteFileExitCode(ByVal exitCode As enmExitCode, ByVal filePath As String)
    On Error Resume Next
    Dim n As Integer
    n = FreeFile
    Open filePath For Output As #n
    Print #n, CStr(exitCode)
    Close #n
End Sub

Private Sub ShowMessage(ByVal message As String, ByVal hConsole As Long, ByVal cmdAttached As Boolean, Optional typeMessage As enmTypeMessage)
    Dim charsWritten As Long
    If cmdAttached Then
        WriteConsole hConsole, message & vbCrLf, Len(message & vbCrLf), charsWritten, ByVal 0&
    Else
        Select Case typeMessage
            Case TypeInformation
                MsgBox message, vbInformation + vbOKOnly
            Case TypeCritical
                MsgBox message, vbCritical + vbOKOnly
            Case TypeExclamation
                MsgBox message, vbExclamation + vbOKOnly
        End Select
    End If
End Sub

Public Function GetProgressBarColor(ByVal prgHwnd As Long) As Long
    Dim hTheme As Long
    Dim color As Long
    Dim result As Long
    
    'Abre el tema para el ProgressBar
    hTheme = OpenThemeData(prgHwnd, "Progress")
    
    If hTheme <> 0 Then
        'Obtiene el color del ProgressBar (parte de la barra, estado normal)
        result = GetThemeColor(hTheme, PBFS_BAR, PBS_NORMAL, TMT_FILLCOLOR, color)
        'Cierra el tema
        Call CloseThemeData(hTheme)
        
        If result = 0 Then
            'Devuelve el color si se obtuvo correctamente
            GetProgressBarColor = color
            Exit Function
        End If
    End If
    
    'En caso de error, devolver un color por defecto (verde)
    GetProgressBarColor = RGB(0, 128, 0)
End Function


Private Function ParseCommandLine(commandLine As String) As String()
    Dim inQuotes As Boolean
    Dim i As Integer
    Dim currentChar As String
    Dim currentArg As String
    Dim args() As String
    ReDim args(0 To 0)
    
    inQuotes = False
    currentArg = ""
    
    For i = 1 To Len(commandLine)
        currentChar = Mid(commandLine, i, 1)
        
        Select Case currentChar
            Case """"
                inQuotes = Not inQuotes
                currentArg = currentArg & currentChar
            Case " "
                If Not inQuotes Then
                    If currentArg <> "" Then
                        args(UBound(args)) = currentArg
                        ReDim Preserve args(0 To UBound(args) + 1)
                        currentArg = ""
                    End If
                Else
                    currentArg = currentArg & currentChar
                End If
            Case Else
                currentArg = currentArg & currentChar
        End Select
    Next i
    
    If currentArg <> "" Then
        args(UBound(args)) = currentArg
    End If

    ParseCommandLine = args
End Function

Private Function RemoveQuotes(value As String) As String
    'Remueve las comillas de un valor si están presentes
    RemoveQuotes = Replace(value, """", "")
End Function

Private Function AppendToArray(arr() As String, value As String) As String()
    Dim arrLength As Integer
    If IsEmpty(arr) Then
        arrLength = 0
    Else
        arrLength = UBound(arr) + 1
    End If
    ReDim Preserve arr(arrLength)
    arr(arrLength) = value
    AppendToArray = arr
End Function

'************************
'* PUBLICS DECLARATIONS *
'************************
Public Sub MessageError(ByVal objErr As ErrObject, Optional ByVal typeMessage As enmTypeMessage)
    Dim strMessage As String
    '---
    strMessage = "The following error occurred in the process.." & vbCrLf & vbCrLf & _
             "Nro. Error: " & objErr.Number & vbCrLf & _
             "Description: " & objErr.Description & vbCrLf & _
             "Source: " & objErr.Source
    '---
    Select Case typeMessage
        Case TypeInformation
            MsgBox strMessage, vbInformation + vbOKOnly
        Case TypeCritical
            MsgBox strMessage, vbCritical + vbOKOnly
        Case TypeExclamation
            MsgBox strMessage, vbExclamation + vbOKOnly
    End Select
End Sub

Public Sub ProgressBackColor(ByVal prgHwnd As Long, Optional ByVal prgBackColor As Long, Optional ByVal prgBarColor As Long)
    If prgHwnd <> 0 Then
        'BACKCOLOR
        If prgBackColor = 0 Then prgBackColor = GetSysColor(COLOR_3DFACE)
        SendMessage prgHwnd, PBM_SETBKCOLOR, 0, prgBackColor
        'BARCOLOR
        If prgBarColor = 0 Then prgBarColor = GetProgressBarColor(prgHwnd)
        SendMessage prgHwnd, PBM_SETBARCOLOR, 0, prgBarColor
    End If
End Sub

Public Sub UnZipFile(ByVal pathZip As String, pathUnZip As String)
    On Error GoTo RoutinError
    Dim objZipFile As Object, objUnZipFolder As Object
    '---
    Set objShell = New Shell 'CreateObject("Shell.Application")
    '--
    Set objZipFile = objShell.NameSpace(pathZip)
    Set objUnZipFolder = objShell.NameSpace(pathUnZip)
    
    If objZipFile Is Nothing Then
        MsgBox "Failed to access the ZIP file path.", vbExclamation + vbOKOnly
        GoTo Finish
    Else
        If GetFileExtension(pathZip) <> "zip" Then
            MsgBox "Downloaded file is not a ZIP file type, it cannot be unzipped.", vbExclamation + vbOKOnly
            GoTo Finish
        End If
    End If
    If objUnZipFolder Is Nothing Then
        'If MsgBox("Destination folder path does not exist." & vbCrLf & vbCrLf & "Do you want to create it?", vbQuestion + vbYesNo) = vbYes Then
            Call MkDir(pathUnZip)
        'Else
        '    MsgBox "Decompression of the file was canceled by the user.", vbInformation + vbOKOnly
        '    Exit Sub
        'End If
    End If
    '--
    Set objUnZipFolder = objShell.NameSpace(pathUnZip)
    objUnZipFolder.CopyHere objZipFile.Items, 4
    '--
Finish:
    Set objShell = Nothing
    Set objZipFile = Nothing
    Set objUnZipFolder = Nothing
    '---
    Exit Sub
RoutinError:
    Call MessageError(Err)
End Sub

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

Public Function GetFileNameFromURL(ByVal url As String, Optional withoutExtension As Boolean = False) As String
    Dim fileName As String
    Dim lastSlashPos    As Integer
    Dim queryStringPos  As Integer
    Dim dotPos          As Integer
    
    'Buscar la última barra diagonal en la URL
    lastSlashPos = InStrRev(url, "/")
    
    'Verificar si hay una cadena de consulta en la URL
    queryStringPos = InStr(url, "?")
    
    'Extraer el nombre del archivo
    If queryStringPos > 0 Then
        fileName = Mid(url, lastSlashPos + 1, queryStringPos - lastSlashPos - 1)
    Else
        fileName = Mid(url, lastSlashPos + 1)
    End If
    
    If withoutExtension Then
        'Buscar la última posición del punto en el nombre del archivo
        dotPos = InStrRev(fileName, ".")
        
        'Si hay un punto, eliminar la extensión
        If dotPos > 0 Then
            fileName = Left(fileName, dotPos - 1)
        End If
    End If
    ' Retornar el nombre del archivo
    GetFileNameFromURL = fileName
End Function

Public Function GetFileExtension(ByVal fileName As String) As String
    Dim dotPos As Integer
    
    'Buscar la última posición del punto en el nombre del archivo
    dotPos = InStrRev(fileName, ".")
    
    'Si hay un punto, extraer la extensión
    If dotPos > 0 Then
        GetFileExtension = Mid(fileName, dotPos + 1)
    Else
        GetFileExtension = ""
    End If
End Function

