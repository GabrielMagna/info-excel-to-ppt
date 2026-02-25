VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'===========================================================
' FICHA N�2 � Excel -> PPT (Plantilla) � Generador masivo
' Estructura: CONFIG -> INSTANCIAS -> VALIDACIONES -> ALGORITMO
' Codigo creado por el mago Gabriel Magna
' Reglas:
' - Hoja: "Ficha 2"
' - Tabla: "TFicha" (A1:O1 headers)
' - Columna A = "N�" (ID y condici�n de parada: primer vac�o => STOP)
' - Columna O = "Estado" (si = OK => saltar; si vac�o u otro texto => generar)
' - Nombre salida: "Ficha N�2 - {N�} - {Nombre Iniciativa}.pptx" (sobrescribe)
'===========================================================

Public Sub Generar_Fichas_PPT_Ficha2()
    On Error GoTo EH
    
    '================= CONFIG =================
    Const SHEET_NAME As String = "Ficha 2"
    Const TABLE_NAME As String = "TFicha"
    Const HEADER_ROW As Long = 5      ' fila encabezados (en la tabla)
    Const START_ROW As Long = 6       ' <-- CAMBIA A 400, 500, etc. cuando quieras (fila Excel)
    ' - Si SOURCE_WB_PATH = "" => usa ThisWorkbook (donde corre el macro)
    Const SOURCE_WB_PATH As String = " " 'aqui va la ruta del excel para macro
    Const TEMPLATE_PPT_PATH As String = " " 'aqui va la ruta de la ppt plantilla
    Const OUTPUT_FOLDER As String = " " ' aqui va la carpeta donde guardara las ppt
    Const ESTADO_OK As String = "OK"
    Const OUTPUT_PREFIX As String = "Ficha N�2 - "
    
    '================= VALIDACIONES RUTAS =================
    Dim outDir As String
    outDir = Trim$(OUTPUT_FOLDER)
    If Len(outDir) = 0 Then Err.Raise vbObjectError + 10, , "Falta configurar OUTPUT_FOLDER."
    If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir

    If Len(Trim$(TEMPLATE_PPT_PATH)) = 0 Then Err.Raise vbObjectError + 11, , "Falta configurar TEMPLATE_PPT_PATH."
    If Dir$(TEMPLATE_PPT_PATH) = "" Then Err.Raise vbObjectError + 12, , "No se encuentra la plantilla: " & TEMPLATE_PPT_PATH

    If START_ROW <= HEADER_ROW Then Err.Raise vbObjectError + 13, , "START_ROW debe ser mayor que HEADER_ROW."

    '================= INSTANCIAS EXCEL =================
    Dim wbSrc As Workbook, ws As Worksheet, lo As ListObject
    If Len(Trim$(SOURCE_WB_PATH)) = 0 Then
        Set wbSrc = ThisWorkbook
    Else
        Set wbSrc = Workbooks.Open(Filename:=SOURCE_WB_PATH, ReadOnly:=False)
    End If
    Set ws = wbSrc.Worksheets(SHEET_NAME)
    Set lo = ws.ListObjects(TABLE_NAME)

    If lo Is Nothing Then Err.Raise vbObjectError + 14, , "No se encontr� la tabla '" & TABLE_NAME & "' en la hoja '" & SHEET_NAME & "'."
    If lo.DataBodyRange Is Nothing Then
        MsgBox "La tabla '" & TABLE_NAME & "' no tiene filas.", vbInformation
        GoTo CleanExit
    End If
    ' Header -> �ndice col
    Dim dictCols As Object
    Set dictCols = CreateObject("Scripting.Dictionary")
    dictCols.CompareMode = 1 ' vbTextCompare

    Dim c As Long, hdr As String
    For c = 1 To lo.HeaderRowRange.Columns.Count
        hdr = Trim$(CStr(lo.HeaderRowRange.Cells(1, c).Value))
        If Len(hdr) > 0 Then dictCols(hdr) = c
    Next c

    RequireHeader dictCols, "N�"
    RequireHeader dictCols, "Nombre Iniciativa"
    RequireHeader dictCols, "Estado"
    '=======================================================
    ' STEP) ALIAS: Placeholder EXACTO del PPT -> Header EXACTO en Excel
    ' (No tocamos la plantilla. Si en PPT dice "lider" sin tilde, ac� se corrige.)
    '=======================================================
    Dim aliasMap As Object
    Set aliasMap = CreateObject("Scripting.Dictionary")
    aliasMap.CompareMode = 1 ' vbTextCompare

    ' --- Diccionario ---
    aliasMap.Add "{N�}", "N�"
    aliasMap.Add "{Nombre Iniciativa}", "Nombre Iniciativa"
    aliasMap.Add "{L�der de Idea}", "L�der de Idea"
    aliasMap.Add "{Cargo l�der}", "Cargo l�der"
    aliasMap.Add "{Integrante 1}", "Integrante 1"
    aliasMap.Add "{Cargoint1}", "Cargoint1"
    aliasMap.Add "{Integrante 2}", "Integrante 2"
    aliasMap.Add "{Cargoint2}", "Cargoint2"
    aliasMap.Add "{Cuadrante}", "Cuadrante"
    aliasMap.Add "{Descripci�n / Contexto}", "Descripci�n / Contexto"
    aliasMap.Add "{Problem�tica}", "Problem�tica"
    aliasMap.Add "{Soluci�n y Beneficios}", "Soluci�n y Beneficios"
    aliasMap.Add "{�C�mo se aborda el problema hoy en d�a?}", "�C�mo se aborda el problema hoy en d�a?"
    aliasMap.Add "{Informaci�n Adicional}", "Informaci�n Adicional"
    ' Validar que todos los headers referenciados existen en la tabla
    Dim k As Variant
    For Each k In aliasMap.Keys
        RequireHeader dictCols, CStr(aliasMap(k))
    Next k

    '================= INSTANCIA POWERPOINT =================
    Dim pptApp As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    '================= ARRAY + LOOP =================
    Dim dataArr As Variant
    dataArr = lo.DataBodyRange.Value ' [1..nRows, 1..nCols]
    Dim nRows As Long
    nRows = UBound(dataArr, 1)
    Dim firstDataRowExcel As Long
    firstDataRowExcel = lo.DataBodyRange.Row
    Dim startIndex As Long
    startIndex = START_ROW - firstDataRowExcel + 1
    If startIndex < 1 Then startIndex = 1
    Dim colID As Long, colNombre As Long, colEstado As Long
    colID = dictCols("N�")
    colNombre = dictCols("Nombre Iniciativa")
    colEstado = dictCols("Estado")
    Dim i As Long
    Dim idVal As String, nombreVal As String, estadoVal As String
    Dim savePath As String
    Dim pptPres As Object
    Dim replaceDict As Object
    For i = startIndex To nRows
        idVal = Trim$(CStr(dataArr(i, colID)))

        ' STOP al primer ID vac�o
        If Len(idVal) = 0 Then Exit For
        estadoVal = Trim$(CStr(dataArr(i, colEstado)))
        If UCase$(estadoVal) = ESTADO_OK Then GoTo NextI
        nombreVal = Trim$(CStr(dataArr(i, colNombre)))
        ' Abrir plantilla como nueva presentaci�n
        Set pptPres = pptApp.Presentations.Open(Filename:=TEMPLATE_PPT_PATH, WithWindow:=msoFalse)
        ' Construir diccionario {placeholderPPT} -> valor desde la columna (header Excel)
        Set replaceDict = CreateObject("Scripting.Dictionary")
        replaceDict.CompareMode = 1 ' vbTextCompare
        For Each k In aliasMap.Keys
            hdr = CStr(aliasMap(k))
            replaceDict.Add CStr(k), CStr(dataArr(i, dictCols(hdr)))
        Next k
        ' Reemplazar
        ReplacePlaceholdersInPresentation pptPres, replaceDict
        ' Guardar (sobrescribe si existe)
        savePath = outDir & SafeFileName(OUTPUT_PREFIX & idVal & " - " & nombreVal) & ".pptx"
        If Dir$(savePath) <> "" Then Kill savePath
        pptPres.SaveAs savePath
        pptPres.Close
        Set pptPres = Nothing
        ' Marcar OK en array
        dataArr(i, colEstado) = ESTADO_OK

NextI:
    Next i
    ' Volcar SOLO Estado a Excel
    Dim estadoRange As Range
    Set estadoRange = lo.DataBodyRange.Columns(colEstado)
    Dim outEstado() As Variant
    ReDim outEstado(1 To nRows, 1 To 1)
    For i = 1 To nRows
        outEstado(i, 1) = dataArr(i, colEstado)
    Next i
    estadoRange.Value = outEstado
    MsgBox "Listo. Fichas generadas (desde fila " & START_ROW & ").", vbInformation

CleanExit:
    On Error Resume Next
    If Not pptPres Is Nothing Then pptPres.Close
    If Not pptApp Is Nothing Then pptApp.Quit
    Set pptPres = Nothing
    Set pptApp = Nothing
    Exit Sub

EH:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'================= VALIDACIONES =================
Private Sub RequireHeader(ByVal dictCols As Object, ByVal headerName As String)
    If Not dictCols.Exists(headerName) Then
        Err.Raise vbObjectError + 200, , "Falta el encabezado '" & headerName & "' en la tabla."
    End If
End Sub

'================= REEMPLAZO (TextRange.Replace) =================
Private Sub ReplacePlaceholdersInPresentation(ByVal pres As Object, ByVal dict As Object)
    Dim sld As Object, shp As Object
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            ReplaceInShape shp, dict
        Next shp
    Next sld
End Sub

Private Sub ReplaceInShape(ByVal shp As Object, ByVal dict As Object)
    On Error Resume Next

    Dim i As Long, r As Long, c As Long, k As Variant

    ' Grupo
    If shp.Type = 6 Then ' msoGroup
        For i = 1 To shp.GroupItems.Count
            ReplaceInShape shp.GroupItems.Item(i), dict
        Next i
        Exit Sub
    End If

    ' Tabla
    If shp.HasTable Then
        With shp.Table
            For r = 1 To .Rows.Count
                For c = 1 To .Columns.Count
                    For Each k In dict.Keys
                        .Cell(r, c).Shape.TextFrame.TextRange.Replace _
                            FindWhat:=CStr(k), ReplaceWhat:=CStr(dict(k)), _
                            After:=0, WholeWords:=0, MatchCase:=0
                    Next k
                Next c
            Next r
        End With
        Exit Sub
    End If

    ' Texto normal
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            For Each k In dict.Keys
                shp.TextFrame.TextRange.Replace _
                    FindWhat:=CStr(k), ReplaceWhat:=CStr(dict(k)), _
                    After:=0, WholeWords:=0, MatchCase:=0
            Next k
        End If
    End If

    On Error GoTo 0
End Sub

'================= SAFE FILE NAME =================
Private Function SafeFileName(ByVal s As String, Optional ByVal maxLen As Long = 120) As String
    Dim bad As Variant, i As Long

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")

    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), " ")
    Next i

    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    If Len(s) = 0 Then s = "Sin Nombre"
    Select Case UCase$(s)
        Case "CON", "PRN", "AUX", "NUL"
            s = s & "_"
    End Select

    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = " ")
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "Sin Nombre"

    If maxLen > 0 And Len(s) > maxLen Then s = Left$(s, maxLen)

    SafeFileName = s
End Function

