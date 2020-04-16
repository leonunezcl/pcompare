Attribute VB_Name = "mVarios"
Option Explicit

Public gbMatchCase As Integer
Public gbWholeWord As Integer
Public gsFindText As String
Public gbLastPos As Integer

Private gsBlackKeywords As String
Public gsBlackKeywords2 As String
Private gsBlueKeyWords As String

Public gsInforme As String
Public gsLastPath As String

'opciones de analisis
Private Type eOptAnalisis
    Value As Integer
End Type

Public Ana_Archivo() As eOptAnalisis
Public Ana_General() As eOptAnalisis
Public Ana_Variables() As eOptAnalisis
Public Ana_Rutinas() As eOptAnalisis

'opciones de configurar para los archivos
Private Type eAnaArchivos
    Nomenclatura As String
    Clase As String
End Type
Public glbAnaArchivos() As eAnaArchivos

'opciones de configurar para los controles
Private Type eAnaControles
    Nomenclatura As String
    Clase As String
End Type
Public glbAnaControles() As eAnaControles

'tipos de variables
Private Type eAnaTipoVariables
    Nomenclatura As String
    TipoVar As String
End Type
Public glbAnaTipoVariables() As eAnaTipoVariables

'tipos de datos
Private Type eAnaAmbitoDatos
    Ambito As String
    Nomenclatura As String
End Type
Public glbAmbitoDatos() As eAnaAmbitoDatos

Public glbLinXArch As Integer
Public glbLarVar As Integer
Public glbLinXRuti As Integer
Public glbMaxNumParam As Integer

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFaceName As String * 33
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'subraya todo aquello que fue analizado y que se detecto no usado
Public Sub ColorizeAnalisisVB(RTF As RichTextBox)

    Dim sBuffer As String
    Dim nI As Long
    Dim nJ As Long
    Dim sTmpWord As String
    Dim nStartPos As Long
    Dim nSelLen As Long
    Dim nWordPos As Long
            
    sBuffer = RTF.Text
    sTmpWord = ""
    With RTF
        For nI = 1 To Len(sBuffer)
            Select Case Mid$(sBuffer, nI, 1)
                Case "A" To "Z", "a" To "z", "_"
                    If sTmpWord = "" Then nStartPos = nI
                    sTmpWord = sTmpWord & Mid$(sBuffer, nI, 1)
                
                Case Chr$(34)
                    nSelLen = 1
                    For nJ = 1 To 9999999
                        If Mid$(sBuffer, nI + 1, 1) = Chr$(34) Then
                            nI = nI + 2
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                
                Case Chr$(39)
                    .SelStart = nI - 1
                    nSelLen = 0
                    For nJ = 1 To 9999999
                        If Mid$(sBuffer, nI, 2) = vbCrLf Then
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                    .SelLength = nSelLen
                    .SelColor = RGB(0, 127, 0)
                
                Case Else
                    If Not (Len(sTmpWord) = 0) Then
                        .SelStart = nStartPos - 1
                        .SelLength = Len(sTmpWord)
                        nWordPos = InStr(1, gsBlackKeywords2, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelStrikeThru = True
                            .SelBold = True
                            .SelColor = RGB(255, 0, 0)
                            .SelText = Mid$(gsBlackKeywords2, nWordPos + 1, Len(sTmpWord))
                        End If
                        nWordPos = InStr(1, gsBlueKeyWords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = RGB(0, 0, 127)
                            .SelText = Mid$(gsBlueKeyWords, nWordPos + 1, Len(sTmpWord))
                        End If
                        If UCase$(sTmpWord) = "REM" Then
                            .SelStart = nI - 4
                            .SelLength = 3
                            For nJ = 1 To 9999999
                                If Mid$(sBuffer, nI, 2) = vbCrLf Then
                                    Exit For
                                Else
                                    .SelLength = .SelLength + 1
                                    nI = nI + 1
                                End If
                            Next
                            .SelColor = RGB(0, 127, 0)
                            .SelText = LCase$(.SelText)
                        End If
                    End If
                    sTmpWord = ""
            End Select
        Next
        .SelStart = 0
    End With

End Sub

'genera un archivo .html
Public Function GuardarArchivoHtml(ByVal Archivo As String, ByVal Titulo As String) As Boolean

    On Local Error GoTo ErrorGuardarArchivoHtml
    
    Dim ret As Boolean
    Dim nFreeFile As Long
    
    ret = True
    
    nFreeFile = FreeFile
    
    Open Archivo For Output As #nFreeFile
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head><title>" & Titulo & "</title></head>"
        Print #nFreeFile, "<body>"
        Print #nFreeFile, gsHtml
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
    
    GoTo SalirGuardarArchivoHtml
    
ErrorGuardarArchivoHtml:
    ret = False
    MsgBox "GuardarArchivoHtml : " & Err & " " & Error$, vbCritical
    Resume SalirGuardarArchivoHtml
    
SalirGuardarArchivoHtml:
    GuardarArchivoHtml = ret
    Err = 0
    
End Function
'convertir el archivo .rtf en archivo .html
Public Function RichToHTML(rtbRichTextBox As RichTextLib.RichTextBox, Optional lngStartPosition As Long, Optional lngEndPosition As Long) As String

    Dim blnBold As Boolean, blnUnderline As Boolean, blnStrikeThru As Boolean
    Dim blnItalic As Boolean, strLastFont As String, lngLastFontColor As Long
    Dim strHTML As String, lngColor As Long, lngRed As Long, lngGreen As Long
    Dim lngBlue As Long, lngCurText As Long, strHex As String, intLastAlignment As Integer

    Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2
    
    'check for lngStartPosition ad lngEndPosition
    
    If IsMissing(lngStartPosition&) Then lngStartPosition& = 0
    If IsMissing(lngEndPosition&) Then lngEndPosition& = Len(rtbRichTextBox.Text)
    
    lngLastFontColor& = -1 'no color

    rtbRichTextBox.Visible = False
    gsCadena = rtbRichTextBox.Text
    
    gsHtml = "<code>"
    DoEvents
   For lngCurText& = lngStartPosition& To lngEndPosition&
       rtbRichTextBox.SelStart = lngCurText&
       rtbRichTextBox.SelLength = 1
   
          If intLastAlignment% <> rtbRichTextBox.SelAlignment Then
             intLastAlignment% = rtbRichTextBox.SelAlignment
              
                Select Case rtbRichTextBox.SelAlignment
                   Case AlignLeft: gsHtml = gsHtml & "<p align=left>"
                   Case AlignRight: gsHtml = gsHtml & "<p align=right>"
                   Case AlignCenter: gsHtml = gsHtml & "<p align=center>"
                End Select
                
          End If
   
          If blnBold <> rtbRichTextBox.SelBold Then
               If rtbRichTextBox.SelBold = True Then
                 gsHtml = gsHtml & "<b>"
               Else
                 gsHtml = gsHtml & "</b>"
               End If
             blnBold = rtbRichTextBox.SelBold
          End If

          If blnUnderline <> rtbRichTextBox.SelUnderline Then
               If rtbRichTextBox.SelUnderline = True Then
                 gsHtml = gsHtml & "<u>"
               Else
                 gsHtml = gsHtml & "</u>"
               End If
             blnUnderline = rtbRichTextBox.SelUnderline
          End If
   

          If blnItalic <> rtbRichTextBox.SelItalic Then
               If rtbRichTextBox.SelItalic = True Then
                 gsHtml = gsHtml & "<i>"
               Else
                 gsHtml = gsHtml & "</i>"
               End If
             blnItalic = rtbRichTextBox.SelItalic
          End If


          If blnStrikeThru <> rtbRichTextBox.SelStrikeThru Then
               If rtbRichTextBox.SelStrikeThru = True Then
                 gsHtml = gsHtml & "<s>"
               Else
                 gsHtml = gsHtml & "</s>"
               End If
             blnStrikeThru = rtbRichTextBox.SelStrikeThru
          End If

         If strLastFont$ <> rtbRichTextBox.SelFontName Then
            strLastFont$ = rtbRichTextBox.SelFontName
            gsHtml = gsHtml + "<font face=""" & strLastFont$ & """>"
         End If

         If lngLastFontColor& <> rtbRichTextBox.SelColor Then
            lngLastFontColor& = rtbRichTextBox.SelColor
            
            ''Get hexidecimal value of color
            strHex$ = Hex(rtbRichTextBox.SelColor)
            strHex$ = String$(6 - Len(strHex$), "0") & strHex$
            strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
            
            gsHtml = gsHtml + "<font color=#" & strHex$ & ">"
        End If
         
        On Error Resume Next
        
        If Asc(Mid$(gsCadena, lngCurText + 1, 1)) <> 13 Then
            gsHtml = gsHtml + rtbRichTextBox.SelText
        Else
            gsHtml = gsHtml + rtbRichTextBox.SelText & "<br>"
        End If
            
   Next lngCurText&
    gsHtml = gsHtml & "</code>"
    rtbRichTextBox.Visible = True
RichToHTML = gsHtml

End Function

Public Sub ColorizeVB(RTF As RichTextBox)
    ' #VBIDEUtils#************************************************************
    ' * Programmer Name : Waty Thierry
    ' * Web Site : http://www.vbdiamond.com
    ' * E-Mail :
    ' * Date : 30/10/98
    ' * Time : 14:47
    ' * Module Name : Colorize_Module
    ' * Module Filename : Colorize.bas
    ' * Procedure Name : ColorizeVB
    ' * Parameters :
    ' * rtf As RichTextBox
    ' **********************************************************************
    ' * Comments : Colorize in black, blue, green the VB keywords
    ' *
    ' *
    ' **********************************************************************
    
    Dim sBuffer As String
    Dim nI As Long
    Dim nJ As Long
    Dim sTmpWord As String
    Dim nStartPos As Long
    Dim nSelLen As Long
    Dim nWordPos As Long
            
    sBuffer = RTF.Text
    sTmpWord = ""
    With RTF
        For nI = 1 To Len(sBuffer)
            Select Case Mid$(sBuffer, nI, 1)
                Case "A" To "Z", "a" To "z", "_"
                    If sTmpWord = "" Then nStartPos = nI
                    sTmpWord = sTmpWord & Mid$(sBuffer, nI, 1)
                
                Case Chr$(34)
                    nSelLen = 1
                    For nJ = 1 To 9999999
                        If Mid$(sBuffer, nI + 1, 1) = Chr$(34) Then
                            nI = nI + 2
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                
                Case Chr$(39)
                    .SelStart = nI - 1
                    nSelLen = 0
                    For nJ = 1 To 9999999
                        If Mid$(sBuffer, nI, 2) = vbCrLf Then
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                    .SelLength = nSelLen
                    .SelColor = RGB(0, 127, 0)
                
                Case Else
                    If Not (Len(sTmpWord) = 0) Then
                        .SelStart = nStartPos - 1
                        .SelLength = Len(sTmpWord)
                        nWordPos = InStr(1, gsBlackKeywords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = RGB(0, 0, 0)
                            .SelText = Mid$(gsBlackKeywords, nWordPos + 1, Len(sTmpWord))
                        End If
                        nWordPos = InStr(1, gsBlueKeyWords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = RGB(0, 0, 127)
                            .SelText = Mid$(gsBlueKeyWords, nWordPos + 1, Len(sTmpWord))
                        End If
                        If UCase$(sTmpWord) = "REM" Then
                            .SelStart = nI - 4
                            .SelLength = 3
                            For nJ = 1 To 9999999
                                If Mid$(sBuffer, nI, 2) = vbCrLf Then
                                    Exit For
                                Else
                                    .SelLength = .SelLength + 1
                                    nI = nI + 1
                                End If
                            Next
                            .SelColor = RGB(0, 127, 0)
                            .SelText = LCase$(.SelText)
                        End If
                    End If
                    sTmpWord = ""
            End Select
        Next
        .SelStart = 0
    End With

End Sub
Public Sub InitColorize()
' **********************************************************************
' * Comments : Initialize the VB keywords
' *
' *
' **********************************************************************

    gsBlackKeywords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear*Collection*Command*Cos*CreateObject*CurDir*DateAdd*DateDiff*DatePart*DateSerial*DateValue*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime*FileLen*Fix*Format*FV*GetAllSettings*GetAttr*GetObject*GetSetting*Hex*Hide*Hour*InputBox*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty*IsError*IsMissing*IsNull*IsNumeric*IsObject*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType*WeekDay*Width*Year*"
    gsBlueKeyWords = "*#Const*#Else*#ElseIf*#End If*#If*Alias*Alias*And*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Name*New*Next*Not*Object*On*Open*Option*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*RSet*Seek*Select*Set*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*Nothing*To*Friend*"

End Sub
Public Function ContarTipoDependencias(ByVal Tipo As eTipoDepencia, Proyecto As eProyecto) As Integer

    Dim k As Integer
    
    Dim ret As Integer
    
    ret = 0
    
    For k = 1 To UBound(Proyecto.aDepencias)
        If Proyecto.aDepencias(k).Tipo = Tipo Then
            ret = ret + 1
        End If
    Next k
    
    ContarTipoDependencias = ret
    
End Function

Public Function ContarTipoRutinas(ByVal Indice As Integer, ByVal Tipo As eTipoRutinas, Proyecto As eProyecto) As Integer

    Dim r As Integer
    
    Dim ret As Integer
    
    ret = 0
    
    For r = 1 To UBound(Proyecto.aArchivos(Indice).aRutinas)
        If Proyecto.aArchivos(Indice).aRutinas(r).Tipo = Tipo Then
            ret = ret + 1
        End If
    Next r
    
    ContarTipoRutinas = ret
    
End Function
Public Function ContarTiposDeArchivos(ByVal Tipo As eTipoArchivo, Proyecto As eProyecto) As Integer

    Dim k As Integer
    
    Dim ret As Integer
    
    ret = 0
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = Tipo Then
            ret = ret + 1
        End If
    Next k
    
    ContarTiposDeArchivos = ret
    
End Function

'muestra las propiedades del archivo seleccionado en el proyecto.
Public Sub PropiedadesArchivo(treeProyecto As Node, Proyecto As eProyecto)

    Dim Nodo As Node
    Dim k As Integer
    Dim Compo As Boolean
    
    Set Nodo = treeProyecto
    
    Compo = False
    
    If InStr(UCase$(Nodo.Text), UCase$("tlb")) Then
        Compo = True
    ElseIf InStr(UCase$(Nodo.Text), UCase$("dll")) Then
        Compo = True
    ElseIf InStr(UCase$(Nodo.Text), UCase$("ocx")) Then
        Compo = True
    ElseIf InStr(UCase$(Nodo.Text), UCase$("olb")) Then
        Compo = True
    End If
    
    'no es componente es un archivo del proyecto
    If Not Compo Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Descripcion = Nodo.Text Then
                Call ShowProperties(Proyecto.aArchivos(k).PathFisico, frmMain.hWnd)
            End If
        Next k
    Else
        Call ShowProperties(Nodo.Text, frmMain.hWnd)
    End If
    
    Set Nodo = Nothing
    
End Sub

Public Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Public Sub Copiar(ByVal hWnd As Long)

    Dim ret As Long
    
    ret = SendMessage(hWnd, WM_COPY, 0, 0)
    
End Sub

Public Function Confirma(ByVal Msg As String) As Integer
    Confirma = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
End Function

Public Sub CargaRutinas(ByVal frm As Form, ByVal Tipo As eTipoRutinas, Proyecto As eProyecto)

    Dim k As Integer
    Dim Itmx As ListItem
    Dim j As Integer
    Dim r As Integer
    
    Call Hourglass(frm.hWnd, True)
    
    j = 1
    For k = 1 To UBound(Proyecto.aArchivos)
'        MsgBox Proyecto.aArchivos(k).Nombre
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 1, 1
                    Set Itmx = frm.lview.ListItems(j)
                    Itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    Itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    Itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 2, 2
                    Set Itmx = frm.lview.ListItems(j)
                    Itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    Itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    Itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 4, 4
                    Set Itmx = frm.lview.ListItems(j)
                    Itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    Itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    Itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 3, 3
                    Set Itmx = frm.lview.ListItems(j)
                    Itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    Itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    Itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        End If
    Next k
    
    Call Hourglass(frm.hWnd, False)
    
    Set Itmx = Nothing
    
End Sub

'busca una
Public Function MyInstr(ByVal Search As String, ByVal What As String) As Boolean
            
    Dim StringArray() As String
    Dim SearchLen As Integer
    Dim k As Integer
    Dim p As Integer
    Dim c As Integer
    Dim Buffer As String
    Dim ret As Boolean
    Dim Chars As String

    ret = False
    p = 1
    c = 0
    Buffer = Search

    If Search = "" Then                     'viene en blanco
        GoTo Salir
    'ElseIf InStr(Search, What) = 0 Then     'hay ocurrencia de alguna substring
    '    GoTo Salir
    End If

Volver:
    Chars = ""
    For k = 1 To Len(Buffer)
        Select Case Mid$(Buffer, k, 1)
            Case "+", "-", "*", "/", ".", ",", "&", " ", "@", "#", "%"
                c = c + 1
                ReDim Preserve StringArray(c)
                StringArray(c) = Trim$(Chars)
                Buffer = Mid$(Buffer, k + 1)
                GoTo Volver
            Case "[", "]", "{", "}", ";", "!", "^", ":"
                c = c + 1
                ReDim Preserve StringArray(c)
                StringArray(c) = Trim$(Chars)
                Buffer = Mid$(Buffer, k + 1)
                GoTo Volver
            Case "$", "(", ")", "=", "\", "<", ">"
                c = c + 1
                ReDim Preserve StringArray(c)
                StringArray(c) = Trim$(Chars)
                Buffer = Mid$(Buffer, k + 1)
                GoTo Volver
            Case Else
                Chars = Chars & Mid$(Buffer, k, 1)
        End Select
    Next k

    c = c + 1
    ReDim Preserve StringArray(c)
    StringArray(c) = Trim$(Chars)

    'validar que no existan caracteres basic
    Select Case Right$(What, 1)
        Case "!", "@", "#", "$", "%", "&"
            What = Left$(What, Len(What) - 1)
    End Select
    
'    ahora ciclar x todas las cadenas encontradas
    For k = 1 To UBound(StringArray())
        If LCase$(StringArray(k)) = LCase$(What) Then
            ret = True
            Exit For
        End If
    Next k
    
Salir:
    MyInstr = ret
    
End Function

Public Sub SelTodo()

    On Local Error Resume Next
    
    'frmMain.txtRutina.SelStart = 0
    'frmMain.txtRutina.SelLength = Len(frmMain.txtRutina.Text)
    'frmMain.txtRutina.SetFocus
    
    Err = 0
    
End Sub


