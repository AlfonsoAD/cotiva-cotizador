Attribute VB_Name = "modSecurity"
'==============================================================================
' modSecurity · Protección del libro y de las hojas
'
' Aviso honesto: la protección de Excel evita cambios accidentales y desalienta
' al usuario común, pero NO es cifrado. Lo que protege el código es la
' contraseña del proyecto VBA (se pone a mano) y, para distribuir, la firma
' digital.
'
' UserInterfaceOnly:=True deja que las MACROS escriban en hojas protegidas,
' pero ese permiso se pierde al cerrar el archivo; por eso ProtectAll se
' vuelve a ejecutar cada vez que se abre el cotizador.
'==============================================================================
Option Explicit

' Intenta quitar la protección de estructura del libro.
' Regresa True si el libro quedó desprotegido (o ya lo estaba).
Public Function TryUnprotectWorkbook() As Boolean
    On Error Resume Next                         ' Si la clave no coincide, no truena aquí...
    ThisWorkbook.Unprotect APP_PASSWORD
    On Error GoTo 0
    TryUnprotectWorkbook = Not ThisWorkbook.ProtectStructure   ' ...lo revisamos aquí
End Function

' Quita la protección de una hoja (sin error si ya estaba desprotegida).
Public Sub UnlockSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect APP_PASSWORD
    On Error GoTo 0
End Sub

' Protege una hoja: celdas bloqueadas e imágenes fijas, pero las macros pueden escribir.
Public Sub LockSheet(ByVal ws As Worksheet)
    UnlockSheet ws
    ws.Protect Password:=APP_PASSWORD, DrawingObjects:=True, Contents:=True, _
               UserInterfaceOnly:=True
End Sub

' Protege todo el libro. Regresa False si la clave no coincide.
Public Function ProtectAll() As Boolean
    Dim sheetName As Variant

    If Not TryUnprotectWorkbook() Then
        MsgBox Fmt(MSG_WRONG_PASSWORD), vbCritical, TTL_WRONG_PASSWORD
        Exit Function
    End If

    ' Hojas de datos: protegidas y (opcionalmente) invisibles
    For Each sheetName In Array(SHEET_QUOTES, SHEET_ITEMS, SHEET_CLIENTS)
        If SheetExists(CStr(sheetName)) Then
            LockSheet ThisWorkbook.Worksheets(CStr(sheetName))
            If HIDE_DATA_SHEETS Then ThisWorkbook.Worksheets(CStr(sheetName)).Visible = xlSheetVeryHidden
        End If
    Next sheetName

    ' Hojas visibles: protegidas (solo las celdas desbloqueadas se pueden editar)
    For Each sheetName In Array(SHEET_TEMPLATE, SHEET_COMPANY, SHEET_HOME)
        If SheetExists(CStr(sheetName)) Then LockSheet ThisWorkbook.Worksheets(CStr(sheetName))
    Next sheetName

    ' Estructura: no se pueden agregar, borrar, renombrar ni mostrar hojas
    ThisWorkbook.Protect Password:=APP_PASSWORD, Structure:=True
    ProtectAll = True
End Function

' [Administrador] Pide la clave y deja todo visible y editable.
Public Sub AdminUnlock()
    Dim answer As String, sheetName As Variant

    answer = InputBox(MSG_ADMIN_PROMPT, TTL_ADMIN)
    If Len(answer) = 0 Then Exit Sub
    If answer <> APP_PASSWORD Then
        MsgBox MSG_ADMIN_WRONG, vbCritical, TTL_ADMIN
        Exit Sub
    End If

    If Not TryUnprotectWorkbook() Then
        MsgBox Fmt(MSG_WRONG_PASSWORD), vbCritical, TTL_WRONG_PASSWORD
        Exit Sub
    End If

    For Each sheetName In Array(SHEET_QUOTES, SHEET_ITEMS, SHEET_CLIENTS, SHEET_TEMPLATE, SHEET_COMPANY, SHEET_HOME)
        If SheetExists(CStr(sheetName)) Then
            ThisWorkbook.Worksheets(CStr(sheetName)).Visible = xlSheetVisible
            UnlockSheet ThisWorkbook.Worksheets(CStr(sheetName))
        End If
    Next sheetName
    MsgBox Fmt(MSG_ADMIN_UNLOCKED), vbInformation, TTL_ADMIN
End Sub

' [Administrador] Vuelve a proteger todo.
Public Sub AdminLock()
    If ProtectAll() Then MsgBox MSG_ADMIN_LOCKED, vbInformation, TTL_ADMIN
End Sub
