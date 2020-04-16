Attribute VB_Name = "MMain"
Option Explicit

Private Arr_Paso() As String
Public glbPBackup As String
Public gbRelease As Boolean
Private cTLI As TypeLibInfo
Public cRegistro As New cRegistry
Private Pathxproyecto As String
Private KeyRegistro As String
Private PathRegistro As String
Private sGUID As String
Private sArchivo As String
Private nLinea As Integer
Public gsHtml As String
Public gsCadena As String
Private REF_DLL As Integer
Private REF_OCX As Integer
Private REF_RES As Integer

Private MayorV As Variant 'As Integer
Private MenorV As Variant 'As Integer
Private P1 As Integer
Private P2 As Integer
Private LineaPaso As String
    
Private FreeSub As Long
Private FreeCodigo As Long
Private StartRutinas As Boolean
Private StartHeader As Boolean
Private StartGeneral As Boolean
Private StartTypes As Boolean
Private StartEnum As Boolean
Private EndGeneral As Boolean
Private EndHeader As Boolean
Private LineaOrigen As String
Private LastTipoRead As Integer
Public gbInicio As Boolean
Public glbSelArchivos As Boolean

'variables que acumulan los tipos analizados
Private ge As Integer
Private k As Integer
Private r As Integer
Private i As Integer
Private c As Integer
Private t As Integer
Private v As Integer
Private e As Integer
Private a As Integer
Private s As Integer
Private f As Integer
Private prop As Integer
Private ap As Integer
Private vr As Integer
Private vpro As Integer
Private aru As Integer
Private apro As Integer
Private ca As Integer
Private even As Integer
Private NumeroDeLineas As Integer

'acumuladores para los tipos
Private spri As Integer
Private spub As Integer

Private fpri As Integer
Private fpub As Integer

Private cpri As Integer
Private cpub As Integer

Private epri As Integer
Private epub As Integer

Private tpri As Integer
Private tpub As Integer

Private apri As Integer
Private apub As Integer

Private vpri As Integer
Private vpub As Integer

Private nFreeFile As Long

'DECLARACION DE LOS TIPOS QUE SON ANALIZADOS
Private Procedimiento As String
Private Funcion As String
Private Constante As String
Private Tipo As String
Private Variable As String
Private Enumeracion As String
Private Arreglo As String
Private Propiedad As String
Private Evento As String

'DECLARACION DE LAS LLAVES DE LOS NODOS
Private PRIFUN As Integer
Private PUBFUN As Integer
Private PRISUB As Integer
Private PUBSUB As Integer
Private PROC As Integer
Private Func As Integer
Private api As Integer
Private Cons As Integer
Private TYPO As Integer
Private TYPOCH As Integer
Private VARY As Integer
Private ENUME As Integer
Private ENUMECH As Integer
Private ARRAYY As Integer
Private VARYPROC As Integer
Private VARYPROP As Integer
Private NPROP As Integer
Private NEVENTO As Integer
Private Linea As String

Private bSub As Boolean
Private bSubPub As Boolean
Private bSubPri As Boolean
Private bEndSub As Boolean
Private bEndProp As Boolean

'DECLARACION DE LOS ICONOS DEL ARBOL PRINCIPAL
Private bFun As Boolean
Private bFunPub As Boolean
Private bFunPri As Boolean
Private bApi As Boolean
Private bCon As Boolean
Private bTipo As Boolean
Private bVariables As Boolean
Private bEnumeracion As Boolean
Private bArray As Boolean
Private bPropiedades As Boolean
Private bEventos As Boolean

Public gsTempPath As String
Public gsRutina As String
Private Sub AcumuladoresParciales(oTrv As TreeView, xproyecto As eProyecto)

    Dim ConVariables As Integer
    Dim ConVrutinas As Integer
    Dim ConRutinas As Integer
    Dim irutina As Integer
    Dim ivar As Integer
    Dim Cantidad As Integer
    
    'acumuladores parciales
                    
    Cantidad = 0
    
    If UBound(xproyecto.aArchivos(k).aVariables) > 0 Then
        
        For ivar = 1 To UBound(xproyecto.aArchivos(k).aTipoVariable)
            Cantidad = Cantidad + xproyecto.aArchivos(k).aTipoVariable(ivar).Cantidad
        Next ivar
        
        xproyecto.aArchivos(k).nVariables = xproyecto.aArchivos(k).nVariables + Cantidad
        
        'MsgBox xproyecto.aArchivos(k).ObjectName & "-" & xproyecto.aArchivos(k).nVariables
                
    End If
                    
    If bCon Then xproyecto.aArchivos(k).nConstantes = c - 1
    If bEnumeracion Then xproyecto.aArchivos(k).nEnumeraciones = e - 1
    If bArray Then xproyecto.aArchivos(k).nArray = a - 1
    If bSub Then xproyecto.aArchivos(k).nRutinas = r - 1
    If bTipo Then xproyecto.aArchivos(k).nTipos = t - 1
    If bSub Then xproyecto.aArchivos(k).nTipoSub = s - 1
    If bFun Then xproyecto.aArchivos(k).nTipoFun = f - 1
    If bApi Then xproyecto.aArchivos(k).nTipoApi = ap - 1
    If bPropiedades Then xproyecto.aArchivos(k).nPropiedades = prop - 1
    If bEventos Then xproyecto.aArchivos(k).nEventos = even - 1
    
    xproyecto.aArchivos(k).TotalLineas = xproyecto.aArchivos(k).NumeroDeLineas - _
                                      xproyecto.aArchivos(k).NumeroDeLineasComentario - _
                                      xproyecto.aArchivos(k).NumeroDeLineasEnBlanco
    
End Sub

Private Sub AcumularTotalesParciales(xTotalesProyecto As eTotalesProyecto, xproyecto As eProyecto, ByVal k, ByVal apri, ByVal apub, ByVal cpri, ByVal cpub, _
                                     ByVal epri, ByVal epub, ByVal fpri, ByVal fpub, _
                                     ByVal spri, ByVal spub, ByVal tpri, ByVal tpub, _
                                     ByVal vpri, ByVal vpub)

    'acumular totales
    xTotalesProyecto.TotalVariables = xTotalesProyecto.TotalVariables + xproyecto.aArchivos(k).nVariables
    xTotalesProyecto.TotalConstantes = xTotalesProyecto.TotalConstantes + xproyecto.aArchivos(k).nConstantes
    xTotalesProyecto.TotalEnumeraciones = xTotalesProyecto.TotalEnumeraciones + xproyecto.aArchivos(k).nEnumeraciones
    xTotalesProyecto.TotalArray = xTotalesProyecto.TotalArray + xproyecto.aArchivos(k).nArray
    xTotalesProyecto.TotalTipos = xTotalesProyecto.TotalTipos + xproyecto.aArchivos(k).nTipos
    xTotalesProyecto.TotalSubs = xTotalesProyecto.TotalSubs + xproyecto.aArchivos(k).nTipoSub
    xTotalesProyecto.TotalFunciones = xTotalesProyecto.TotalFunciones + xproyecto.aArchivos(k).nTipoFun
    xTotalesProyecto.TotalApi = xTotalesProyecto.TotalApi + xproyecto.aArchivos(k).nTipoApi
    
    xTotalesProyecto.TotalLineasDeCodigo = xTotalesProyecto.TotalLineasDeCodigo + xproyecto.aArchivos(k).NumeroDeLineas
    xTotalesProyecto.TotalLineasDeComentarios = xTotalesProyecto.TotalLineasDeComentarios + xproyecto.aArchivos(k).NumeroDeLineasComentario
    xTotalesProyecto.TotalLineasEnBlancos = xTotalesProyecto.TotalLineasEnBlancos + xproyecto.aArchivos(k).NumeroDeLineasEnBlanco
    
    xTotalesProyecto.TotalPropiedades = xTotalesProyecto.TotalPropiedades + xproyecto.aArchivos(k).nPropiedades
    xTotalesProyecto.TotalEventos = xTotalesProyecto.TotalEventos + xproyecto.aArchivos(k).nEventos
    
    xTotalesProyecto.TotalArrayPrivadas = xTotalesProyecto.TotalArrayPrivadas + (apri - 1)
    xTotalesProyecto.TotalArrayPublicas = xTotalesProyecto.TotalArrayPublicas + (apub - 1)
    
    xTotalesProyecto.TotalConstantesPrivadas = xTotalesProyecto.TotalConstantesPrivadas + (cpri - 1)
    xTotalesProyecto.TotalConstantesPublicas = xTotalesProyecto.TotalConstantesPublicas + (cpub - 1)
    
    xTotalesProyecto.TotalEnumeracionesPrivadas = xTotalesProyecto.TotalEnumeracionesPrivadas + (epri - 1)
    xTotalesProyecto.TotalEnumeracionesPublicas = xTotalesProyecto.TotalEnumeracionesPublicas + (epub - 1)
    
    xTotalesProyecto.TotalFuncionesPrivadas = xTotalesProyecto.TotalFuncionesPrivadas + (fpri - 1)
    xTotalesProyecto.TotalFuncionesPublicas = xTotalesProyecto.TotalFuncionesPublicas + (fpub - 1)
    
    xTotalesProyecto.TotalSubsPrivadas = xTotalesProyecto.TotalSubsPrivadas + (spri - 1)
    xTotalesProyecto.TotalSubsPublicas = xTotalesProyecto.TotalSubsPublicas + (spub - 1)
    
    xTotalesProyecto.TotalTiposPrivadas = xTotalesProyecto.TotalTiposPrivadas + (tpri - 1)
    xTotalesProyecto.TotalTiposPublicas = xTotalesProyecto.TotalTiposPublicas + (tpub - 1)
    
    xTotalesProyecto.TotalVariablesPrivadas = xTotalesProyecto.TotalVariablesPrivadas + (vpri - 1)
    xTotalesProyecto.TotalVariablesPublicas = xTotalesProyecto.TotalVariablesPublicas + (vpub - 1)
        
    xTotalesProyecto.TotalMiembrosPrivados = xTotalesProyecto.TotalMiembrosPrivados + xproyecto.aArchivos(k).MiembrosPrivados
    xTotalesProyecto.TotalMiembrosPublicos = xTotalesProyecto.TotalMiembrosPublicos + xproyecto.aArchivos(k).MiembrosPublicos
    
End Sub

'agrega el archivo de xproyecto a estructura
Private Sub AgregaArchivoDexproyecto(oTrv As TreeView, xproyecto As eProyecto, k As Integer, ByVal Archivo As String, _
                                    ByVal Tipo As eTipoArchivo, ByVal KeyArchivo As String)

    ReDim Preserve xproyecto.aArchivos(k)
                    
    'CHEQUEAR \
    If PathArchivo(Archivo) = "" Then
        xproyecto.aArchivos(k).Nombre = Archivo
        xproyecto.aArchivos(k).PathFisico = Pathxproyecto & Archivo
    Else
        xproyecto.aArchivos(k).Nombre = Mid$(Archivo, InStr(Archivo, "\") + 1)
        xproyecto.aArchivos(k).PathFisico = Pathxproyecto & Archivo
    End If
    
    xproyecto.aArchivos(k).TipoDeArchivo = Tipo
    
    If Tipo = TIPO_ARCHIVO_FRM Then
        xproyecto.aArchivos(k).KeyNodeFrm = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_BAS Then
        xproyecto.aArchivos(k).KeyNodeBas = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_CLS Then
        xproyecto.aArchivos(k).KeyNodeCls = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_OCX Then
        xproyecto.aArchivos(k).KeyNodeKtl = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_PAG Then
        xproyecto.aArchivos(k).KeyNodePag = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_REL Then
        xproyecto.aArchivos(k).KeyNodeRel = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_DSR Then
        xproyecto.aArchivos(k).KeyNodeDsr = KeyArchivo & k
    End If
    
    xproyecto.aArchivos(k).FILETIME = VBGetFileTime(xproyecto.aArchivos(k).PathFisico)
    xproyecto.aArchivos(k).Explorar = True
    
    k = k + 1
                    
End Sub

'agrega los componentes al arbol de xproyecto
Private Sub AgregaComponentes(oTrv As TreeView, xproyecto As eProyecto, d As Integer, ByVal Linea As String)

    On Local Error Resume Next
    
    'BUSCAR MAYOR
    P1 = 0: P2 = 0
    P1 = InStr(1, Linea, "#")
    P2 = InStr(P1 + 1, Linea, "#") - 1
    MayorV = Mid$(Linea, P1 + 1, P2 - P1)
    
    'BUSCAR MENOR
    P1 = InStr(P2, Linea, ";") - 1
    MenorV = Mid$(Linea, P2 + 2, P1 - P2)
    If Right$(MenorV, 1) = ";" Then MenorV = Left$(MenorV, Len(MenorV) - 1)

    sGUID = Left$(Linea, InStr(1, Linea, "}"))
    sGUID = Mid$(sGUID, 8)
    
    If InStr(1, MayorV, ".") Then
        MenorV = Mid$(MayorV, InStr(1, MayorV, ".") + 1)
        MayorV = Left$(MayorV, InStr(1, MayorV, ".") - 1)
    End If
    
    Set cTLI = TLI.TypeLibInfoFromRegistry(sGUID, Val(MayorV), Val(MenorV), 0)
    
    If Err <> 0 Then
        MsgBox LoadResString(C_ERROR_DEPENDENCIA) & vbNewLine & sArchivo, vbCritical
    Else
        ReDim Preserve xproyecto.aDepencias(d)
        
        xproyecto.aDepencias(d).Archivo = cTLI.ContainingFile
        xproyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
        xproyecto.aDepencias(d).HelpString = cTLI.HelpString
        xproyecto.aDepencias(d).HelpFile = cTLI.HelpFile
        xproyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
        xproyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
        xproyecto.aDepencias(d).GUID = cTLI.GUID
        xproyecto.aDepencias(d).Tipo = TIPO_OCX
        xproyecto.aDepencias(d).FileSize = VBGetFileSize(xproyecto.aDepencias(d).Archivo)
        xproyecto.aDepencias(d).FILETIME = VBGetFileTime(xproyecto.aDepencias(d).Archivo)
        xproyecto.aDepencias(d).KeyNode = "REFOCX" & REF_OCX
        REF_OCX = REF_OCX + 1
        d = d + 1
    End If
    
    Err = 0
                    
End Sub
'agrega la referencias
Private Sub AgregaReferencias(oTrv As TreeView, xproyecto As eProyecto, d As Integer, ByVal Linea As String)

    On Local Error Resume Next
    
    'BUSCAR MAYOR
    P1 = 0: P2 = 0
    P1 = InStr(1, Linea, "#")
    P2 = InStr(P1 + 1, Linea, "#") - 1
    MayorV = Mid$(Linea, P1 + 1, P2 - P1)
    
    'BUSCAR MENOR
    P1 = InStr(P2 + 2, Linea, "#") - 1
    MenorV = Mid$(Linea, P2 + 2, P1 - P2)
    If Right$(MenorV, 1) = "#" Then
        MenorV = Left$(MenorV, Len(MenorV) - 1)
    End If
    
    KeyRegistro = Mid$(Linea, InStr(1, Linea, "G") + 1)
    KeyRegistro = Left$(KeyRegistro, InStr(1, KeyRegistro, "}"))
                    
    cRegistro.ClassKey = HKEY_CLASSES_ROOT
    cRegistro.ValueType = REG_SZ
    cRegistro.SectionKey = "TypeLib\" & KeyRegistro & "\" & Val(MayorV) & "\" & Val(MenorV) & "\win32"
    sArchivo = cRegistro.Value
    
    If sArchivo = "" Then
        sArchivo = NombreArchivo(Linea, 1)
    End If
            
    Set cTLI = TLI.TypeLibInfoFromRegistry(KeyRegistro, Val(MayorV), Val(MenorV), 0)
    
    If Err.Number <> 0 Then
        Err = 0
        Set cTLI = TLI.TypeLibInfoFromFile(sArchivo)
    
        If Err.Number <> 0 Then
            MsgBox LoadResString(C_ERROR_DEPENDENCIA) & vbNewLine & sArchivo, vbCritical
        Else
            ReDim Preserve xproyecto.aDepencias(d)
        
            xproyecto.aDepencias(d).Archivo = sArchivo
            
            xproyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
            xproyecto.aDepencias(d).HelpString = cTLI.HelpString
            xproyecto.aDepencias(d).HelpFile = cTLI.HelpFile
            xproyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
            xproyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
            xproyecto.aDepencias(d).Tipo = TIPO_DLL
            xproyecto.aDepencias(d).GUID = cTLI.GUID
            xproyecto.aDepencias(d).FileSize = VBGetFileSize(xproyecto.aDepencias(d).Archivo)
            xproyecto.aDepencias(d).FILETIME = VBGetFileTime(xproyecto.aDepencias(d).Archivo)
            xproyecto.aDepencias(d).KeyNode = "REFDLL" & REF_DLL
            REF_DLL = REF_DLL + 1
            d = d + 1
        End If
    Else
        ReDim Preserve xproyecto.aDepencias(d)
        
        If xproyecto.Version > 3 Or xproyecto.Version = 0 Then
            xproyecto.aDepencias(d).Archivo = sArchivo
            xproyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
            xproyecto.aDepencias(d).HelpString = cTLI.HelpString
            xproyecto.aDepencias(d).HelpFile = cTLI.HelpFile
            xproyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
            xproyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
            xproyecto.aDepencias(d).Tipo = TIPO_DLL
            xproyecto.aDepencias(d).GUID = cTLI.GUID
        Else
            xproyecto.aDepencias(d).Archivo = Linea
            xproyecto.aDepencias(d).ContainingFile = Linea
            xproyecto.aDepencias(d).HelpString = ""
            xproyecto.aDepencias(d).HelpFile = 0
            xproyecto.aDepencias(d).MajorVersion = 0
            xproyecto.aDepencias(d).MinorVersion = 0
            xproyecto.aDepencias(d).Tipo = TIPO_DLL
            xproyecto.aDepencias(d).GUID = ""
        End If
        
        xproyecto.aDepencias(d).FileSize = VBGetFileSize(xproyecto.aDepencias(d).Archivo)
        xproyecto.aDepencias(d).FILETIME = VBGetFileTime(xproyecto.aDepencias(d).Archivo)
        xproyecto.aDepencias(d).KeyNode = "REFDLL" & REF_DLL
        REF_DLL = REF_DLL + 1
        d = d + 1
    End If
    
    Err = 0
                
End Sub

'agrega el tipo de funcion segun el archivo que esta siendo procesado
'publica o privada
Private Sub AgregaTipoDeFuncion(oTrv As TreeView, xproyecto As eProyecto, ByVal Publica As Boolean)

    Dim KeyNode As String
    Dim Icono As Integer
    Dim Glosa As String
    
    If Publica Then
        KeyNode = xproyecto.aArchivos(k).KeyNodeFun & "-FPUB" & PUBFUN
        Icono = C_ICONO_PUBLIC_FUNCION
        PUBFUN = PUBFUN + 1
        Glosa = LoadResString(C_PUBLICAS)
    Else
        KeyNode = xproyecto.aArchivos(k).KeyNodeFun & "-FPRI" & PRIFUN
        Icono = C_ICONO_PRIVATE_FUNCION
        PRIFUN = PRIFUN + 1
        Glosa = LoadResString(C_PRIVADAS)
    End If
    
    'agregar el tipo de funcion al arbol
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_FUNC_FRM & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_FUNC_BAS & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_FUNC_CLS & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_FUNC_CTL & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_FUNC_PAG & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_FUNC_DSR & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    End If
    
End Sub
'agrega el tipo de sub al arbol
Private Sub AgregaTipoDeSub(oTrv As TreeView, xproyecto As eProyecto, ByVal Publica As Boolean)

    Dim KeyNode As String
    Dim Icono As Integer
    Dim Glosa As String
    
    If Publica Then
        KeyNode = xproyecto.aArchivos(k).KeyNodeSub & "-SPUB" & PUBSUB
        Icono = C_ICONO_PUBLIC_SUB
        PUBSUB = PUBSUB + 1
        Glosa = LoadResString(C_PUBLICAS)
    Else
        KeyNode = xproyecto.aArchivos(k).KeyNodeSub & "-SPRI" & PRISUB
        Icono = C_ICONO_PRIVATE_SUB
        PRISUB = PRISUB + 1
        Glosa = LoadResString(C_PRIVADAS)
    End If
    
    'agregar el tipo de funcion al arbol
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_SUB_FRM & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_SUB_BAS & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_SUB_CLS & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_SUB_CTL & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_SUB_PAG & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_SUB_DSR & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    End If
    
End Sub
'ANALIZAR APIS
Private Sub AnalizaApi(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreVar As String
    Dim Libreria As String
    
    Funcion = NombreX(Linea)
    ReDim Preserve xproyecto.aArchivos(k).aApis(ap)
    xproyecto.aArchivos(k).aApis(ap).Nombre = Funcion
    xproyecto.aArchivos(k).aApis(ap).Estado = NOCHEQUEADO
    
    If Left$(Funcion, 23) = LoadResString(C_PUBLIC_DECLARE_FUNCTION) Then
        Funcion = Mid$(Funcion, 25)
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        xproyecto.aArchivos(k).aApis(ap).Publica = True
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
    ElseIf Left$(Funcion, 24) = LoadResString(C_PRIVATE_DECLARE_FUNCTION) Then
        Funcion = Mid$(Funcion, 26)
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        xproyecto.aArchivos(k).aApis(ap).Publica = False
        xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
    ElseIf Left$(Funcion, 16) = LoadResString(C_DECLARE_FUNCTION) Then
        Funcion = Mid$(Funcion, 18)
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        xproyecto.aArchivos(k).aApis(ap).Publica = True
        xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
    ElseIf Left$(Funcion, 18) = LoadResString(C_PUBLIC_DECLARE_SUB) Then
        Funcion = Mid$(Funcion, 20)
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        xproyecto.aArchivos(k).aApis(ap).Publica = True
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
    ElseIf Left$(Funcion, 19) = LoadResString(C_PRIVATE_DECLARE_SUB) Then
        Funcion = Mid$(Funcion, 21)
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        xproyecto.aArchivos(k).aApis(ap).Publica = False
        xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
    Else
        Funcion = Mid$(Funcion, 13)
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        xproyecto.aArchivos(k).aApis(ap).Publica = True
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
    End If
    
    'comprobar si api esta declarada al viejo estilo basic
    If Right$(xproyecto.aArchivos(k).aApis(ap).NombreVariable, 1) = "%" Then
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(xproyecto.aArchivos(k).aApis(ap).NombreVariable, Len(xproyecto.aArchivos(k).aApis(ap).NombreVariable) - 1)
        xproyecto.aArchivos(k).aApis(ap).BasicOldStyle = True
    ElseIf Right$(xproyecto.aArchivos(k).aApis(ap).NombreVariable, 1) = "&" Then
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(xproyecto.aArchivos(k).aApis(ap).NombreVariable, Len(xproyecto.aArchivos(k).aApis(ap).NombreVariable) - 1)
        xproyecto.aArchivos(k).aApis(ap).BasicOldStyle = True
    ElseIf Right$(xproyecto.aArchivos(k).aApis(ap).NombreVariable, 1) = "$" Then
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(xproyecto.aArchivos(k).aApis(ap).NombreVariable, Len(xproyecto.aArchivos(k).aApis(ap).NombreVariable) - 1)
        xproyecto.aArchivos(k).aApis(ap).BasicOldStyle = True
    ElseIf Right$(xproyecto.aArchivos(k).aApis(ap).NombreVariable, 1) = "#" Then
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(xproyecto.aArchivos(k).aApis(ap).NombreVariable, Len(xproyecto.aArchivos(k).aApis(ap).NombreVariable) - 1)
        xproyecto.aArchivos(k).aApis(ap).BasicOldStyle = True
    ElseIf Right$(xproyecto.aArchivos(k).aApis(ap).NombreVariable, 1) = "@" Then
        xproyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(xproyecto.aArchivos(k).aApis(ap).NombreVariable, Len(xproyecto.aArchivos(k).aApis(ap).NombreVariable) - 1)
        xproyecto.aArchivos(k).aApis(ap).BasicOldStyle = True
    End If
    
    'agregar nodo principal
    If Not bApi Then
        If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call oTrv.Nodes.Add(C_KEY_FRM & k, tvwChild, "FAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            xproyecto.aArchivos(k).KeyNodeApi = "FAPROC" & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call oTrv.Nodes.Add(C_KEY_BAS & k, tvwChild, "BAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            xproyecto.aArchivos(k).KeyNodeApi = "BAPROC" & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call oTrv.Nodes.Add(C_KEY_CLS & k, tvwChild, "CAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            xproyecto.aArchivos(k).KeyNodeApi = "CAPROC" & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call oTrv.Nodes.Add(C_KEY_CTL & k, tvwChild, "KAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            xproyecto.aArchivos(k).KeyNodeApi = "KAPROC" & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call oTrv.Nodes.Add(C_KEY_PAG & k, tvwChild, "PAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            xproyecto.aArchivos(k).KeyNodeApi = "PAPROC" & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            Call oTrv.Nodes.Add(C_KEY_DSR & k, tvwChild, "DAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            xproyecto.aArchivos(k).KeyNodeApi = "DAPROC" & k
        End If
        bApi = True
    End If

    'determinar a que libreria pertenece
    NombreVar = xproyecto.aArchivos(k).aApis(ap).NombreVariable
    Libreria = Mid$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") + 5)
    Libreria = Left$(Libreria, InStr(1, Libreria, Chr$(34)) - 1)
    
    On Error Resume Next
    
    'agregar la libreria al arbol de librerias
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add("FAPROC" & k, tvwChild, Libreria & "FAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add("BAPROC" & k, tvwChild, Libreria & "BAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add("CAPROC" & k, tvwChild, Libreria & "CAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add("KAPROC" & k, tvwChild, Libreria & "KAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add("PAPROC" & k, tvwChild, Libreria & "PAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add("DAPROC" & k, tvwChild, Libreria & "DAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    End If
    
    'agregar la funcion segun la libreria
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(Libreria & "FAPROC" & k, tvwChild, "FAPI" & api, NombreVar, C_ICONO_API, C_ICONO_API)
        xproyecto.aArchivos(k).aApis(ap).KeyNode = "FAPI" & api
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(Libreria & "BAPROC" & k, tvwChild, "BAPI" & api, NombreVar, C_ICONO_API, C_ICONO_API)
        xproyecto.aArchivos(k).aApis(ap).KeyNode = "BAPI" & api
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(Libreria & "CAPROC" & k, tvwChild, "CAPI" & api, NombreVar, C_ICONO_API, C_ICONO_API)
        xproyecto.aArchivos(k).aApis(ap).KeyNode = "CAPI" & api
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(Libreria & "KAPROC" & k, tvwChild, "KAPI" & api, NombreVar, C_ICONO_API, C_ICONO_API)
        xproyecto.aArchivos(k).aApis(ap).KeyNode = "KAPI" & api
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(Libreria & "PAPROC" & k, tvwChild, "PAPI" & api, NombreVar, C_ICONO_API, C_ICONO_API)
        xproyecto.aArchivos(k).aApis(ap).KeyNode = "PAPI" & api
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(Libreria & "DAPROC" & k, tvwChild, "DAPI" & api, NombreVar, C_ICONO_API, C_ICONO_API)
        xproyecto.aArchivos(k).aApis(ap).KeyNode = "DAPI" & api
    End If
    
    xproyecto.aArchivos(k).aApis(ap).Linea = nLinea
    
    Err = 0
    
    api = api + 1
    ap = ap + 1
    nLinea = nLinea + 1
    
End Sub

'ANALIZA ARREGLOS
Private Sub AnalizaArray(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreArray As String
            
    Linea = NombreX(Linea)
    
    If Left$(Linea, 3) = Trim$(LoadResString(C_DIM)) Then
        Variable = Mid$(Variable, 5)
        ReDim Preserve xproyecto.aArchivos(k).aArray(a)
        xproyecto.aArchivos(k).aArray(a).Publica = False
        xproyecto.aArchivos(k).aArray(a).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
        apub = apub + 1
    ElseIf Left$(Linea, 7) = Trim$(LoadResString(C_PRIVATE)) Then
        Variable = Mid$(Variable, 9)
        ReDim Preserve xproyecto.aArchivos(k).aArray(a)
        xproyecto.aArchivos(k).aArray(a).Publica = False
        xproyecto.aArchivos(k).aArray(a).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
        apri = apri + 1
    ElseIf Left$(Linea, 6) = Trim$(LoadResString(C_PUBLIC)) Then
        Variable = Mid$(Variable, 8)
        ReDim Preserve xproyecto.aArchivos(k).aArray(a)
        xproyecto.aArchivos(k).aArray(a).Publica = True
        xproyecto.aArchivos(k).aArray(a).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
        apub = apub + 1
    ElseIf Left$(Linea, 6) = Trim$(LoadResString(C_GLOBAL)) Then
        ReDim Preserve xproyecto.aArchivos(k).aArray(a)
        Variable = Mid$(Variable, 8)
        xproyecto.aArchivos(k).aArray(a).Publica = True
        xproyecto.aArchivos(k).aArray(a).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
        apub = apub + 1
    Else
        Exit Sub
    End If
    
    xproyecto.aArchivos(k).aArray(a).Nombre = Variable
    Variable = Left$(Variable, InStr(Variable, "(") - 1)
    xproyecto.aArchivos(k).aArray(a).NombreVariable = Variable
        
    'agregar el nodo de array
    If Not bArray Then
        If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call oTrv.Nodes.Add(C_KEY_FRM & k, tvwChild, C_ARR_FRM & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            xproyecto.aArchivos(k).KeyNodeArr = C_ARR_FRM & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call oTrv.Nodes.Add(C_KEY_BAS & k, tvwChild, C_ARR_BAS & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            xproyecto.aArchivos(k).KeyNodeArr = C_ARR_BAS & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call oTrv.Nodes.Add(C_KEY_CLS & k, tvwChild, C_ARR_CLS & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            xproyecto.aArchivos(k).KeyNodeArr = C_ARR_CLS & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call oTrv.Nodes.Add(C_KEY_CTL & k, tvwChild, C_ARR_CTL & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            xproyecto.aArchivos(k).KeyNodeArr = C_ARR_CTL & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call oTrv.Nodes.Add(C_KEY_PAG & k, tvwChild, C_ARR_PAG & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            xproyecto.aArchivos(k).KeyNodeArr = C_ARR_PAG & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            Call oTrv.Nodes.Add(C_KEY_DSR & k, tvwChild, C_ARR_DSR & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            xproyecto.aArchivos(k).KeyNodeArr = C_ARR_DSR & k
        End If
        bArray = True
    End If
    
    'agregar el array al nodo de arrays
    NombreArray = xproyecto.aArchivos(k).aArray(a).NombreVariable
    
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_ARR_FRM & k, tvwChild, "FARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
        xproyecto.aArchivos(k).aArray(a).KeyNode = "FARR" & ARRAYY
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_ARR_BAS & k, tvwChild, "BARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
        xproyecto.aArchivos(k).aArray(a).KeyNode = "BARR" & ARRAYY
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_ARR_CLS & k, tvwChild, "CARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
        xproyecto.aArchivos(k).aArray(a).KeyNode = "CARR" & ARRAYY
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_ARR_CTL & k, tvwChild, "KARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
        xproyecto.aArchivos(k).aArray(a).KeyNode = "KARR" & ARRAYY
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_ARR_PAG & k, tvwChild, "PARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
        xproyecto.aArchivos(k).aArray(a).KeyNode = "PARR" & ARRAYY
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_ARR_DSR & k, tvwChild, "DARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
        xproyecto.aArchivos(k).aArray(a).KeyNode = "DARR" & ARRAYY
    End If
    
    xproyecto.aArchivos(k).aArray(a).Linea = nLinea
    
    nLinea = nLinea + 1
    a = a + 1
    ARRAYY = ARRAYY + 1
                    
End Sub

Private Sub AnalizaDim(oTrv As TreeView, xproyecto As eProyecto)

    Dim sVariable As String
    Dim TipoVb As String
    Dim Inicio As Integer
    Dim Inicio2 As Integer
    Dim Fin As Boolean
    Dim nTipoVar As Integer
    Dim Predefinido As Boolean
    Dim NombreEnum As String
        
    If Left$(Linea, 1) <> "'" Then  'COMENTARIO ?
    
        Linea = NombreX(Linea)
        
        'VARIABLES GLOBALES A NIVEL GENERAL O LOCAL
        'NO LAS VARIABLES INTERIORES
        Variable = Linea
        sVariable = Variable
        Inicio = 1
        Inicio2 = 0
        Fin = True
        
        If Left$(Variable, 12) <> LoadResString(C_PRIVATE_ENUM) And Left$(Variable, 11) <> LoadResString(C_PUBLIC_ENUM) Then
            Do  'CICLAR HASTA QUE NO HAYA MAS A CHEQUEAR ?
                If InStr(1, sVariable, ",") <> 0 Then
                    Variable = Left$(sVariable, InStr(1, sVariable, ",") - 1)
                    Inicio = InStr(1, sVariable, ",") + 1
                    sVariable = Trim$(Mid$(sVariable, Inicio))
                    Fin = False
                Else
                    Variable = sVariable
                    Fin = True
                End If
                
                Variable = Trim$(Variable)
                
                If InStr(Variable, "(") = 0 Then    'ARRAY ?
                    If StartRutinas Then    'variables a nivel de rutinas
                        ReDim Preserve xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr)
                        xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre = Variable
                        xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                        xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).TipoVb = TipoVb
                        xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Predefinido = Predefinido
                        xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Estado = NOCHEQUEADO
                        xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).BasicOldStyle = BasicOldStyle(Variable)
                        xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Linea = nLinea
                        
                        'tipo de variable string,byte,currency,etc
                        Call ProcesarTipoDeVariable(oTrv, xproyecto, Variable)
                    ElseIf Not StartRutinas Then    'VARIABLES GLOBALES A NIVEL DEL TIPO DE ARCHIVO
                        ReDim Preserve xproyecto.aArchivos(k).aVariables(v)
                        xproyecto.aArchivos(k).aVariables(v).Nombre = Variable
                        xproyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                        xproyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                        xproyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                        xproyecto.aArchivos(k).aVariables(v).Estado = NOCHEQUEADO
                        xproyecto.aArchivos(k).aVariables(v).BasicOldStyle = BasicOldStyle(Variable)
                        xproyecto.aArchivos(k).aVariables(v).Linea = nLinea
                        
                        'tipo de variable string,byte,currency,etc
                        Call ProcesarTipoDeVariable(oTrv, xproyecto, Variable)
                    End If
                    
                    If Left$(Variable, 3) = Trim$(LoadResString(C_DIM)) Then
                        Variable = Mid$(Variable, 5)
                        Variable = Trim$(Variable)
                        
                        If StartRutinas Then    'variables a nivel de rutinas
                            If InStr(Variable, LoadResString(C_AS)) Then
                                xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Variable)
                            End If
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).TipoVb = TipoVb
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Predefinido = Predefinido
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Estado = NOCHEQUEADO
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Publica = False
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Linea = nLinea
                            xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
                        ElseIf Not StartRutinas Then    'VARIABLES GLOBALES A NIVEL DEL TIPO DE ARCHIVO
                            If InStr(Variable, LoadResString(C_AS)) Then
                                xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                            End If
                            xproyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            xproyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                            xproyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                            xproyecto.aArchivos(k).aVariables(v).Estado = NOCHEQUEADO
                            xproyecto.aArchivos(k).aVariables(v).Publica = True
                            xproyecto.aArchivos(k).aVariables(v).UsaDim = True
                            xproyecto.aArchivos(k).aVariables(v).Linea = nLinea
                            xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
                            vpub = vpub + 1
                        End If
                    ElseIf Left$(Variable, 6) = Trim$(LoadResString(C_STATIC)) Then
                        Variable = Mid$(Variable, 8)
                        
                        If StartRutinas Then    'variables a nivel de rutinas
                            If InStr(Variable, LoadResString(C_AS)) Then
                                xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Variable)
                            End If
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).TipoVb = TipoVb
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Predefinido = Predefinido
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Estado = NOCHEQUEADO
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Publica = False
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Linea = nLinea
                            xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
                            
                        ElseIf Not StartRutinas Then    'VARIABLES GLOBALES A NIVEL DEL TIPO DE ARCHIVO
                            If InStr(Variable, LoadResString(C_AS)) Then
                                xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                            End If
                            xproyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            xproyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                            xproyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                            xproyecto.aArchivos(k).aVariables(v).Estado = NOCHEQUEADO
                            xproyecto.aArchivos(k).aVariables(v).Publica = True
                            xproyecto.aArchivos(k).aVariables(v).Linea = nLinea
                            xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
                            vpub = vpub + 1
                        End If
                    ElseIf Left$(Variable, 7) = Trim$(LoadResString(C_PRIVATE)) Then
                        Variable = Mid$(Variable, 9)
                        
                        If InStr(Variable, LoadResString(C_AS)) Then
                            xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                        Else
                            xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                        End If
                        xproyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                        xproyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                        xproyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                        xproyecto.aArchivos(k).aVariables(v).Estado = NOCHEQUEADO
                        xproyecto.aArchivos(k).aVariables(v).Publica = False
                        xproyecto.aArchivos(k).aVariables(v).Linea = nLinea
                        xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
                        vpri = vpri + 1
                    ElseIf Left$(Variable, 6) = Trim$(LoadResString(C_PUBLIC)) Then
                        Variable = Mid$(Variable, 8)
                        
                        If InStr(Variable, LoadResString(C_AS)) Then
                            xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                        Else
                            xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                        End If
                        xproyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                        xproyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                        xproyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                        xproyecto.aArchivos(k).aVariables(v).Estado = NOCHEQUEADO
                        xproyecto.aArchivos(k).aVariables(v).Publica = True
                        xproyecto.aArchivos(k).aVariables(v).Linea = nLinea
                        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
                        vpub = vpub + 1
                    ElseIf Left$(Variable, 6) = Trim$(LoadResString(C_GLOBAL)) Then
                        Variable = Mid$(Variable, 8)
                        
                        If InStr(Variable, LoadResString(C_AS)) Then
                            xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                        Else
                            xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                        End If
                        xproyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                        xproyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                        xproyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                        xproyecto.aArchivos(k).aVariables(v).Estado = NOCHEQUEADO
                        xproyecto.aArchivos(k).aVariables(v).Publica = True
                        xproyecto.aArchivos(k).aVariables(v).UsaGlobal = True
                        xproyecto.aArchivos(k).aVariables(v).Linea = nLinea
                        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
                        vpub = vpub + 1
                    Else    'ES SECUENCIA DE , DIM A,
                        Variable = "Dim " & Variable
                        Variable = Mid$(Variable, 5)
                                                                            
                        If StartRutinas Then    'variables a nivel de rutinas
                            If InStr(Variable, LoadResString(C_AS)) Then
                                xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Variable)
                            End If
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).TipoVb = TipoVb
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Predefinido = Predefinido
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Estado = NOCHEQUEADO
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Publica = False
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).BasicOldStyle = BasicOldStyle(Variable)
                            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Linea = nLinea
                            xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
                            
                        ElseIf Not StartRutinas Then    'VARIABLES GLOBALES A NIVEL DEL TIPO DE ARCHIVO
                            If InStr(Variable, LoadResString(C_AS)) Then
                                xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                xproyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                            End If
                            xproyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            xproyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                            xproyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                            xproyecto.aArchivos(k).aVariables(v).Estado = NOCHEQUEADO
                            xproyecto.aArchivos(k).aVariables(v).Publica = True
                            xproyecto.aArchivos(k).aVariables(v).BasicOldStyle = BasicOldStyle(Variable)
                            xproyecto.aArchivos(k).aVariables(v).Linea = nLinea
                            xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
                            vpub = vpub + 1
                        End If
                    End If
                    
                    'agregar icono de variable al arbol del xproyecto
                    Call StartDim(oTrv, xproyecto)
                    
                    'agregar hijo de variable
                    Call StartChildDim(oTrv, xproyecto)
                    
                    'If StartPropiedades Then
                    '    vpro = vpro + 1
                    '    VARYPROP = VARYPROP + 1
                    If Not StartRutinas Then
                        v = v + 1
                        VARY = VARY + 1
                    ElseIf StartRutinas Then
                        vr = vr + 1
                        VARYPROC = VARYPROC + 1
                    End If
                Else
                    Call AnalizaArray(oTrv, xproyecto)
                End If
            Loop Until Fin
            nLinea = nLinea + 1
        Else
            Call AnalizaEnumeracion(oTrv, xproyecto)
        End If
    Else
        xproyecto.aArchivos(k).NumeroDeLineasComentario = xproyecto.aArchivos(k).NumeroDeLineasComentario + 1
    End If
                        
End Sub
'cargar enumeraciones
Private Sub AnalizaEnumeracion(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreEnum As String
    
    Enumeracion = NombreX(Linea)
    
    If Left$(Enumeracion, 13) = LoadResString(C_PRIVATE_ENUM) Then
        Enumeracion = Mid$(Enumeracion, 14)
        ReDim Preserve xproyecto.aArchivos(k).aEnumeraciones(e)
        ReDim xproyecto.aArchivos(k).aEnumeraciones(e).aElementos(0)
        
        xproyecto.aArchivos(k).aEnumeraciones(e).NombreVariable = Enumeracion
        xproyecto.aArchivos(k).aEnumeraciones(e).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).aEnumeraciones(e).Publica = False
        xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
        epri = epri + 1
    ElseIf Left$(Enumeracion, 12) = LoadResString(C_PUBLIC_ENUM) Then
        Enumeracion = Mid$(Enumeracion, 13)
        ReDim Preserve xproyecto.aArchivos(k).aEnumeraciones(e)
        ReDim xproyecto.aArchivos(k).aEnumeraciones(e).aElementos(0)
        
        xproyecto.aArchivos(k).aEnumeraciones(e).NombreVariable = Enumeracion
        xproyecto.aArchivos(k).aEnumeraciones(e).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).aEnumeraciones(e).Publica = True
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
        epub = epub + 1
    ElseIf Left$(Enumeracion, 5) = LoadResString(C_ENUM) Then
        Enumeracion = Mid$(Enumeracion, 6)
        ReDim Preserve xproyecto.aArchivos(k).aEnumeraciones(e)
        ReDim xproyecto.aArchivos(k).aEnumeraciones(e).aElementos(0)
        
        xproyecto.aArchivos(k).aEnumeraciones(e).NombreVariable = Enumeracion
        xproyecto.aArchivos(k).aEnumeraciones(e).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).aEnumeraciones(e).Publica = True
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
        epub = epub + 1
    Else
        Exit Sub
    End If
    
    'para comenzar a guardar los elementos de la enumeracion
    If Not StartEnum Then
        StartEnum = True
    End If
        
    xproyecto.aArchivos(k).aEnumeraciones(e).Nombre = Enumeracion
    
    'agregar enumeracion al arbol del xproyecto
    If Not bEnumeracion Then
        If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call oTrv.Nodes.Add(C_KEY_FRM & k, tvwChild, C_ENUM_FRM & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            xproyecto.aArchivos(k).KeyNodeEnum = C_ENUM_FRM & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call oTrv.Nodes.Add(C_KEY_BAS & k, tvwChild, C_ENUM_BAS & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            xproyecto.aArchivos(k).KeyNodeEnum = C_ENUM_BAS & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call oTrv.Nodes.Add(C_KEY_CLS & k, tvwChild, C_ENUM_CLS & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            xproyecto.aArchivos(k).KeyNodeEnum = C_ENUM_CLS & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call oTrv.Nodes.Add(C_KEY_CTL & k, tvwChild, C_ENUM_CTL & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            xproyecto.aArchivos(k).KeyNodeEnum = C_ENUM_CTL & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call oTrv.Nodes.Add(C_KEY_PAG & k, tvwChild, C_ENUM_PAG & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            xproyecto.aArchivos(k).KeyNodeEnum = C_ENUM_PAG & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            Call oTrv.Nodes.Add(C_KEY_DSR & k, tvwChild, C_ENUM_DSR & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            xproyecto.aArchivos(k).KeyNodeEnum = C_ENUM_DSR & k
        End If
        bEnumeracion = True
    End If
    
    NombreEnum = xproyecto.aArchivos(k).aEnumeraciones(e).NombreVariable
    
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_ENUM_FRM & k, tvwChild, "FENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        xproyecto.aArchivos(k).aEnumeraciones(e).KeyNode = "FENUM" & ENUME
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_ENUM_BAS & k, tvwChild, "BENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        xproyecto.aArchivos(k).aEnumeraciones(e).KeyNode = "BENUM" & ENUME
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_ENUM_CLS & k, tvwChild, "CENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        xproyecto.aArchivos(k).aEnumeraciones(e).KeyNode = "CENUM" & ENUME
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_ENUM_CTL & k, tvwChild, "KENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        xproyecto.aArchivos(k).aEnumeraciones(e).KeyNode = "KENUM" & ENUME
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_ENUM_PAG & k, tvwChild, "PENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        xproyecto.aArchivos(k).aEnumeraciones(e).KeyNode = "PENUM" & ENUME
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_ENUM_DSR & k, tvwChild, "DENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        xproyecto.aArchivos(k).aEnumeraciones(e).KeyNode = "DENUM" & ENUME
    End If
    
    xproyecto.aArchivos(k).aEnumeraciones(e).Linea = nLinea
    
    nLinea = nLinea + 1
    e = e + 1
    ENUME = ENUME + 1
                        
End Sub

'CARGAR EVENTOS ...
Private Sub AnalizaEvento(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreEvento As String
    
    Evento = NombreX(Linea)
    
    If Left$(Evento, 6) = LoadResString(C_EVENTO) Or Left$(Evento, 13) = LoadResString(C_PUBLIC_EVENT) Then
    
        ReDim Preserve xproyecto.aArchivos(k).aEventos(even)
                
        If InStr(1, Evento, "'") = 0 Then
            xproyecto.aArchivos(k).aEventos(even).Nombre = Evento
        Else
            xproyecto.aArchivos(k).aEventos(even).Nombre = Trim$(Left$(Evento, InStr(1, Evento, "'") - 1))
        End If
            
        xproyecto.aArchivos(k).aEventos(even).Estado = NOCHEQUEADO
        
        If Left$(Evento, 6) = LoadResString(C_EVENTO) Then
            xproyecto.aArchivos(k).aEventos(even).Publica = True
            Evento = Mid$(Evento, 7)
        ElseIf Left$(Evento, 13) = LoadResString(C_PUBLIC_EVENT) Then
            xproyecto.aArchivos(k).aEventos(even).Publica = True
            Evento = Mid$(Evento, 14)
        End If
        
        If InStr(1, Evento, "'") Then
            Evento = Trim$(Left$(Evento, InStr(1, Evento, "'") - 1))
        End If
        
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
        
        If InStr(Evento, "(") <> 0 Then
            'If InStr(Evento, C_AS) Then
                xproyecto.aArchivos(k).aEventos(even).NombreVariable = Left$(Evento, InStr(1, Evento, "(") - 1)
            'Else
            '    xproyecto.aArchivos(k).aEventos(even).NombreVariable = Evento
            'End If
        Else
            xproyecto.aArchivos(k).aEventos(even).NombreVariable = Evento
        End If
            
        'agregar evento al arbol del xproyecto
        If Not bEventos Then
            If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                Call oTrv.Nodes.Add(C_KEY_FRM & k, tvwChild, C_EVEN_FRM & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                xproyecto.aArchivos(k).KeyNodeEvento = C_EVEN_FRM & k
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                Call oTrv.Nodes.Add(C_KEY_BAS & k, tvwChild, C_EVEN_BAS & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                xproyecto.aArchivos(k).KeyNodeEvento = C_EVEN_BAS & k
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                Call oTrv.Nodes.Add(C_KEY_CLS & k, tvwChild, C_EVEN_CLS & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                xproyecto.aArchivos(k).KeyNodeEvento = C_EVEN_CLS & k
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                Call oTrv.Nodes.Add(C_KEY_CTL & k, tvwChild, C_EVEN_CTL & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                xproyecto.aArchivos(k).KeyNodeEvento = C_EVEN_CTL & k
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                Call oTrv.Nodes.Add(C_KEY_PAG & k, tvwChild, C_EVEN_PAG & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                xproyecto.aArchivos(k).KeyNodeEvento = C_EVEN_PAG & k
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                Call oTrv.Nodes.Add(C_KEY_DSR & k, tvwChild, C_EVEN_DSR & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                xproyecto.aArchivos(k).KeyNodeEvento = C_EVEN_DSR & k
            End If
            bEventos = True
        End If
                    
        'agregar evento al arbol
        NombreEvento = xproyecto.aArchivos(k).aEventos(even).NombreVariable
        
        If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call oTrv.Nodes.Add(C_EVEN_FRM & k, tvwChild, "FE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
            xproyecto.aArchivos(k).aEventos(even).KeyNode = "FE_EVEN" & NEVENTO
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call oTrv.Nodes.Add(C_EVEN_BAS & k, tvwChild, "BE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
            xproyecto.aArchivos(k).aEventos(even).KeyNode = "BE_EVEN" & NEVENTO
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call oTrv.Nodes.Add(C_EVEN_CLS & k, tvwChild, "CE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
            xproyecto.aArchivos(k).aEventos(even).KeyNode = "CE_EVEN" & NEVENTO
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call oTrv.Nodes.Add(C_EVEN_CTL & k, tvwChild, "KE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
            xproyecto.aArchivos(k).aEventos(even).KeyNode = "KE_EVEN" & NEVENTO
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call oTrv.Nodes.Add(C_EVEN_PAG & k, tvwChild, "PE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
            xproyecto.aArchivos(k).aEventos(even).KeyNode = "PE_EVEN" & NEVENTO
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            Call oTrv.Nodes.Add(C_EVEN_DSR & k, tvwChild, "DE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
            xproyecto.aArchivos(k).aEventos(even).KeyNode = "DE_EVEN" & NEVENTO
        End If
                        
        xproyecto.aArchivos(k).aEventos(even).Linea = nLinea
        
        even = even + 1
        NEVENTO = NEVENTO + 1
        nLinea = nLinea + 1
        
    End If
    
End Sub

'analizar funcion
Private Sub AnalizaFunction(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreFuncion As String
    Dim LineaX As String
    
    LineaX = NombreX(Linea)
    Linea = NombreX(Linea)
    
    If Left$(Linea, 8) = Trim$(LoadResString(C_FUNCTION)) Or Left$(Linea, 15) = Trim$(LoadResString(C_FRIEND_FUNCTION)) Then
        Funcion = Linea
        ReDim Preserve xproyecto.aArchivos(k).aRutinas(r)
        xproyecto.aArchivos(k).aRutinas(r).Nombre = Funcion
        xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
        
        If Left$(Linea, 8) = Trim$(LoadResString(C_FUNCTION)) Then
            Funcion = Mid$(Funcion, 10)
        ElseIf Left$(Linea, 15) = Trim$(LoadResString(C_FRIEND_FUNCTION)) Then
            Funcion = Mid$(Funcion, 17)
        End If
        
        xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Funcion, InStr(1, Funcion, "(") - 1)
        xproyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN
        xproyecto.aArchivos(k).aRutinas(r).Publica = True
        xproyecto.aArchivos(k).aRutinas(r).Linea = nLinea
        xproyecto.aArchivos(k).aRutinas(r).BasicStyle = BasicOldStyle(xproyecto.aArchivos(k).aRutinas(r).NombreRutina)
        ReDim xproyecto.aArchivos(k).aRutinas(r).aVariables(0)
        ReDim xproyecto.aArchivos(k).aRutinas(r).aRVariables(0)
        
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
        
        If Not StartRutinas Then
            StartRutinas = True
            Call StartRastreoRutinas(xproyecto)
        End If
        
        If Not bFun Then Call StartFuncion(oTrv, xproyecto)
            
        NombreFuncion = xproyecto.aArchivos(k).aRutinas(r).NombreRutina
        
        If Not bFunPub Then
            Call AgregaTipoDeFuncion(oTrv, xproyecto, True)
            bFunPub = True
        End If
    
        Call StartChildFuncion(oTrv, xproyecto, NombreFuncion, C_ICONO_PUBLIC_FUNCION, xproyecto.aArchivos(k).KeyNodeFun & "-FPUB" & PUBFUN - 1)
                
        ReDim xproyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
        nLinea = nLinea + 1
        r = r + 1
        Func = Func + 1
        f = f + 1
        fpub = fpub + 1
        vr = 1 'para contar variables rutinas
        
        'chequear si no viene la fun cortada
        Linea = Mid$(Linea, Len(LoadResString(C_FUNCTION)) + 1)
        Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)

        If Linea <> ")" Then
            Call ProcesarParametros(oTrv, xproyecto)
            If Right$(xproyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
                xproyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
            Else
                xproyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
                xproyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Funcion)
            End If
        End If
    ElseIf InStr(Linea, LoadResString(C_DECLARE)) Then
        Call AnalizaApi(oTrv, xproyecto)
    End If
                        
    Linea = LineaX
    
End Sub
'analizar nombre del control
Private Sub AnalizaNombreControl(oTrv As TreeView, xproyecto As eProyecto, xTotalesProyecto As eTotalesProyecto)

    Dim sControl As String
    Dim TipoOcx As Integer
    Dim sNombre As String
    Dim j As Integer
    Dim ChControl As Integer
    Dim FoundCtrl As Boolean
    
    If MyInstr(Linea, LoadResString(C_FORM)) = False Then
        If MyInstr(Linea, Trim$(LoadResString(C_BEGIN))) Then
            FoundCtrl = False
            sControl = Mid$(Linea, InStr(Linea, " ") + 1)
            sControl = Mid$(sControl, InStr(sControl, " ") + 1)
            
            'comprobar si es un array de controles
            For ChControl = 1 To UBound(xproyecto.aArchivos(k).aControles())
                If sControl = xproyecto.aArchivos(k).aControles(ChControl).Nombre Then
                    xproyecto.aArchivos(k).aControles(ChControl).Numero = _
                    xproyecto.aArchivos(k).aControles(ChControl).Numero + 1
                    xproyecto.aArchivos(k).aControles(ChControl).Descripcion = _
                    "(" & sControl & "-" & xproyecto.aArchivos(k).aControles(ChControl).Numero & ")"
                    FoundCtrl = True
                    Exit For
                End If
            Next ChControl
            
            If Not FoundCtrl Then
                sControl = Mid$(Linea, InStr(Linea, " ") + 1)
                sControl = Mid$(sControl, InStr(sControl, " ") + 1)
                
                ReDim Preserve xproyecto.aArchivos(k).aControles(ca)
                xproyecto.aArchivos(k).aControles(ca).Nombre = sControl
                xproyecto.aArchivos(k).aControles(ca).Numero = 1
                xproyecto.aArchivos(k).aControles(ca).Descripcion = sControl
                
                sControl = Mid$(Linea, InStr(Linea, " ") + 1)
                sControl = Left$(sControl, InStr(sControl, " ") - 1)
                
                xproyecto.aArchivos(k).aControles(ca).Clase = sControl
                
                xproyecto.aArchivos(k).nControles = xproyecto.aArchivos(k).nControles + 1
                xTotalesProyecto.TotalControles = xTotalesProyecto.TotalControles + 1
                
                ca = ca + 1
            End If
        End If
    End If
                        
End Sub
'analizar private const
Private Sub AnalizaPrivateConst(oTrv As TreeView, xproyecto As eProyecto)

    Dim Okey As Boolean
    Dim NombreConstante As String
    
    Okey = False
    
    Constante = NombreX(Linea)
    
    If Left$(Constante, 6) = LoadResString(C_CONST) Then Okey = True
    If Left$(Constante, 14) = LoadResString(C_PRIVATE_CONST) Then Okey = True
    If Left$(Constante, 13) = LoadResString(C_PUBLIC_CONST) Then Okey = True
    If Left$(Constante, 13) = LoadResString(C_GLOBAL_CONST) Then Okey = True
    
    If Not Okey Then Exit Sub
        
    ReDim Preserve xproyecto.aArchivos(k).aConstantes(c)
    xproyecto.aArchivos(k).aConstantes(c).Nombre = Constante
        
    If Left$(Constante, 6) = LoadResString(C_CONST) Then
        xproyecto.aArchivos(k).aConstantes(c).Publica = False
        If Not StartRutinas Then
            xproyecto.aArchivos(k).aConstantes(c).UsaPrivate = True
        End If
    ElseIf Left$(Constante, 14) = LoadResString(C_PRIVATE_CONST) Then
        xproyecto.aArchivos(k).aConstantes(c).Publica = False
    ElseIf Left$(Constante, 13) = LoadResString(C_PUBLIC_CONST) Then
        xproyecto.aArchivos(k).aConstantes(c).Publica = True
    ElseIf Left$(Constante, 13) = LoadResString(C_GLOBAL_CONST) Then
        xproyecto.aArchivos(k).aConstantes(c).Publica = True
        If Not StartRutinas Then
            xproyecto.aArchivos(k).aConstantes(c).UsaGlobal = True
        End If
    End If
    
    xproyecto.aArchivos(k).aConstantes(c).Estado = NOCHEQUEADO
    
    If Left$(Constante, 6) = LoadResString(C_CONST) Then
        Constante = Mid$(Constante, 7)
    ElseIf Left$(Constante, 14) = LoadResString(C_PRIVATE_CONST) Then
        Constante = Mid$(Constante, 15)
    ElseIf Left$(Constante, 13) = LoadResString(C_PUBLIC_CONST) Then
        Constante = Mid$(Constante, 14)
    ElseIf Left$(Constante, 13) = LoadResString(C_GLOBAL_CONST) Then
        Constante = Mid$(Constante, 14)
    End If
        
    Constante = Left$(Constante, InStr(1, Constante, "=") - 2)
        
    If InStr(Constante, LoadResString(C_AS)) Then
        Constante = Left$(Constante, InStr(Constante, LoadResString(C_AS)) - 1)
    End If
        
    xproyecto.aArchivos(k).aConstantes(c).NombreVariable = Constante
    xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
    xproyecto.aArchivos(k).aConstantes(c).Linea = nLinea
    
    If Not bCon Then Call StartConstantes(oTrv, xproyecto)
    
    NombreConstante = xproyecto.aArchivos(k).aConstantes(c).NombreVariable
    
    Call StartChildConstante(oTrv, xproyecto, NombreConstante)
    
    nLinea = nLinea + 1
    c = c + 1
    Cons = Cons + 1
    cpri = cpri + 1
                        
End Sub

'analizar private function
Private Sub AnalizaPrivateFunction(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreFuncion As String
    Dim LineaX As String
    
    LineaX = NombreX(Linea)
    Linea = NombreX(Linea)
    
    Funcion = Linea
    ReDim Preserve xproyecto.aArchivos(k).aRutinas(r)
    xproyecto.aArchivos(k).aRutinas(r).Nombre = Funcion
    
    Funcion = Mid$(Funcion, 18)
    xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Funcion, InStr(1, Funcion, "(") - 1)
    xproyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN
    xproyecto.aArchivos(k).aRutinas(r).Publica = False
    xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
    xproyecto.aArchivos(k).aRutinas(r).Linea = nLinea
    xproyecto.aArchivos(k).aRutinas(r).BasicStyle = BasicOldStyle((xproyecto.aArchivos(k).aRutinas(r).NombreRutina))
    xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
    
    ReDim xproyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim xproyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas(xproyecto)
    End If
    
    If Not bFun Then Call StartFuncion(oTrv, xproyecto)
        
    NombreFuncion = xproyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    If Not bFunPri Then
        Call AgregaTipoDeFuncion(oTrv, xproyecto, False)
        bFunPri = True
    End If
    
    Call StartChildFuncion(oTrv, xproyecto, NombreFuncion, C_ICONO_PRIVATE_FUNCION, xproyecto.aArchivos(k).KeyNodeFun & "-FPRI" & PRIFUN - 1)
                            
    ReDim xproyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
    nLinea = nLinea + 1
    r = r + 1
    f = f + 1
    fpri = fpri + 1
    Func = Func + 1
    vr = 1 'para contar variables rutinas
    
    Linea = Mid$(Linea, Len(LoadResString(C_PRIVATE_FUNCTION)) + 1)
    Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)

    If Linea <> ")" Then
        Call ProcesarParametros(oTrv, xproyecto)
        If Right$(xproyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
            xproyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
        Else
            xproyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
            xproyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Funcion)
        End If
    End If
        
    Linea = LineaX
    
End Sub

'analizar propiedad
Private Sub AnalizaPropiedad(oTrv As TreeView, xproyecto As eProyecto, xTotalesProyecto As eTotalesProyecto)

    Dim NombrePropiedad As String
    Dim Privada As Boolean
    Dim Icono As Integer
    Dim LineaX As String
    Dim LineaPaso As String
    
    LineaX = NombreX(Linea)
    LineaPaso = NombreX(Linea)
    
    Propiedad = NombreX(Linea)
    Privada = False
    
    vr = 1 'para contar variables rutinas
    vpro = 1
    
    ReDim Preserve xproyecto.aArchivos(k).aRutinas(r)
                    
    xproyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_PROPIEDAD
    
    If Left$(Propiedad, 21) = LoadResString(C_PROP_PRIVATE_GET) Then
    
        LineaPaso = Mid$(Linea, Len(LoadResString(C_PROP_PRIVATE_GET)) + 1)
        LineaPaso = Mid$(Linea, InStr(1, Linea, "(") + 1)
    
        xproyecto.aArchivos(k).aRutinas(r).Nombre = Propiedad
        xproyecto.aArchivos(k).aRutinas(r).Publica = False
        
        Propiedad = Mid$(Propiedad, 22)
        If InStr(Propiedad, "(") <> 0 Then
            If InStr(Propiedad, LoadResString(C_AS)) Then
                xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Propiedad, InStr(Propiedad, LoadResString(C_AS)) - 3)
            Else
                xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Propiedad, InStr(Propiedad, "(") - 1)
            End If
        Else
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Propiedad
        End If
        xproyecto.aArchivos(k).aRutinas(r).TipoProp = TIPO_GET
        xproyecto.aArchivos(k).aRutinas(r).Publica = False
        xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
        Privada = True
        'ACUMULADORES
        xproyecto.aArchivos(k).nPropertyGet = xproyecto.aArchivos(k).nPropertyGet + 1
        xTotalesProyecto.TotalPropertyGets = xTotalesProyecto.TotalPropertyGets + 1
    ElseIf Left$(Propiedad, 21) = LoadResString(C_PROP_PRIVATE_LET) Then
        
        LineaPaso = Mid$(Linea, Len(LoadResString(C_PROP_PRIVATE_LET)) + 1)
        LineaPaso = Mid$(Linea, InStr(1, Linea, "(") + 1)
        
        xproyecto.aArchivos(k).aRutinas(r).Nombre = Propiedad
        xproyecto.aArchivos(k).aRutinas(r).Publica = False
        xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
        Propiedad = Mid$(Propiedad, 22)
        If InStr(Propiedad, "(") <> 0 Then
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Propiedad
        End If
        xproyecto.aArchivos(k).aRutinas(r).TipoProp = TIPO_LET
        xproyecto.aArchivos(k).aRutinas(r).Publica = False
        Privada = True
        'ACUMULADORES
        xproyecto.aArchivos(k).nPropertyLet = xproyecto.aArchivos(k).nPropertyLet + 1
        xTotalesProyecto.TotalPropertyLets = xTotalesProyecto.TotalPropertyLets + 1
    ElseIf Left$(Propiedad, 21) = LoadResString(C_PROP_PRIVATE_SET) Then
        
        LineaPaso = Mid$(Linea, Len(LoadResString(C_PROP_PRIVATE_SET)) + 1)
        LineaPaso = Mid$(Linea, InStr(1, Linea, "(") + 1)
        
        xproyecto.aArchivos(k).aRutinas(r).Nombre = Propiedad
        xproyecto.aArchivos(k).aRutinas(r).Publica = False
        xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
        Propiedad = Mid$(Propiedad, 22)
        If InStr(Propiedad, "(") <> 0 Then
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Propiedad
        End If
        
        xproyecto.aArchivos(k).aRutinas(r).TipoProp = TIPO_SET
        xproyecto.aArchivos(k).aRutinas(r).Publica = False
        Privada = True
        'ACUMULADORES
        xproyecto.aArchivos(k).nPropertySet = xproyecto.aArchivos(k).nPropertySet + 1
        xTotalesProyecto.TotalPropertySets = xTotalesProyecto.TotalPropertySets + 1
    ElseIf Left$(Propiedad, 20) = LoadResString(C_PROP_PUBLIC_GET) Then
        
        LineaPaso = Mid$(Linea, Len(LoadResString(C_PROP_PUBLIC_GET)) + 1)
        LineaPaso = Mid$(Linea, InStr(1, Linea, "(") + 1)
        
        xproyecto.aArchivos(k).aRutinas(r).Nombre = Propiedad
        xproyecto.aArchivos(k).aRutinas(r).Publica = True
        xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
        Propiedad = Mid$(Propiedad, 21)
        If InStr(Propiedad, "(") <> 0 Then
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Propiedad
        End If
        xproyecto.aArchivos(k).aRutinas(r).TipoProp = TIPO_GET
        xproyecto.aArchivos(k).aRutinas(r).Publica = True
        
        'ACUMULAR
        xproyecto.aArchivos(k).nPropertyGet = xproyecto.aArchivos(k).nPropertyGet + 1
        xTotalesProyecto.TotalPropertyGets = xTotalesProyecto.TotalPropertyGets + 1
    ElseIf Left$(Propiedad, 20) = LoadResString(C_PROP_PUBLIC_LET) Then
        
        LineaPaso = Mid$(Linea, Len(LoadResString(C_PROP_PUBLIC_LET)) + 1)
        LineaPaso = Mid$(Linea, InStr(1, Linea, "(") + 1)
        
        xproyecto.aArchivos(k).aRutinas(r).Nombre = Propiedad
        xproyecto.aArchivos(k).aRutinas(r).Publica = True
        xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
        Propiedad = Mid$(Propiedad, 21)
        If InStr(Propiedad, "(") <> 0 Then
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Propiedad
        End If
        xproyecto.aArchivos(k).aRutinas(r).TipoProp = TIPO_LET
        xproyecto.aArchivos(k).aRutinas(r).Publica = True
        
        'ACUMULAR
        xproyecto.aArchivos(k).nPropertyLet = xproyecto.aArchivos(k).nPropertyLet + 1
        xTotalesProyecto.TotalPropertyLets = xTotalesProyecto.TotalPropertyLets + 1
    ElseIf Left$(Propiedad, 20) = LoadResString(C_PROP_PUBLIC_SET) Then
        
        LineaPaso = Mid$(Linea, Len(LoadResString(C_PROP_PUBLIC_SET)) + 1)
        LineaPaso = Mid$(Linea, InStr(1, Linea, "(") + 1)
        
        xproyecto.aArchivos(k).aRutinas(r).Nombre = Propiedad
        Propiedad = Mid$(Propiedad, 21)
        If InStr(Propiedad, "(") <> 0 Then
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Propiedad
        End If
        
        xproyecto.aArchivos(k).aRutinas(r).Publica = True
        xproyecto.aArchivos(k).aRutinas(r).TipoProp = TIPO_SET
        xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
        'ACUMULAR
        xproyecto.aArchivos(k).nPropertySet = xproyecto.aArchivos(k).nPropertySet + 1
        xTotalesProyecto.TotalPropertySets = xTotalesProyecto.TotalPropertySets + 1
    ElseIf Left$(Propiedad, 19) = "Friend Property Get" Then
        LineaPaso = Mid$(Linea, 21)
        LineaPaso = Mid$(Linea, InStr(1, Linea, "(") + 1)
        
        xproyecto.aArchivos(k).aRutinas(r).Nombre = Propiedad
        xproyecto.aArchivos(k).aRutinas(r).Publica = True
        xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
        Propiedad = Mid$(Propiedad, 21)
        If InStr(Propiedad, "(") <> 0 Then
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Propiedad
        End If
        xproyecto.aArchivos(k).aRutinas(r).TipoProp = TIPO_GET
        xproyecto.aArchivos(k).aRutinas(r).Publica = True
        
        'ACUMULAR
        xproyecto.aArchivos(k).nPropertyGet = xproyecto.aArchivos(k).nPropertyGet + 1
        xTotalesProyecto.TotalPropertyGets = xTotalesProyecto.TotalPropertyGets + 1
    ElseIf Left$(Propiedad, 19) = "Friend Property Let" Then
    
    ElseIf Left$(Propiedad, 19) = "Friend Property Set" Then
    
    End If
    
    'agregar propiedades al arbol del xproyecto
    If Not bPropiedades Then
        If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call oTrv.Nodes.Add(C_KEY_FRM & k, tvwChild, C_PROP_FRM & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            xproyecto.aArchivos(k).KeyNodeProp = C_PROP_FRM & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call oTrv.Nodes.Add(C_KEY_BAS & k, tvwChild, C_PROP_BAS & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            xproyecto.aArchivos(k).KeyNodeProp = C_PROP_BAS & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call oTrv.Nodes.Add(C_KEY_CLS & k, tvwChild, C_PROP_CLS & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            xproyecto.aArchivos(k).KeyNodeProp = C_PROP_CLS & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call oTrv.Nodes.Add(C_KEY_CTL & k, tvwChild, C_PROP_CTL & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            xproyecto.aArchivos(k).KeyNodeProp = C_PROP_CTL & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call oTrv.Nodes.Add(C_KEY_PAG & k, tvwChild, C_PROP_PAG & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            xproyecto.aArchivos(k).KeyNodeProp = C_PROP_PAG & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            Call oTrv.Nodes.Add(C_KEY_DSR & k, tvwChild, C_PROP_DSR & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            xproyecto.aArchivos(k).KeyNodeProp = C_PROP_DSR & k
        End If
        bPropiedades = True
    End If
                
    NombrePropiedad = xproyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    If Privada Then
        Icono = C_ICONO_PROPIEDAD_PRIVADA
        xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
    Else
        Icono = C_ICONO_PROPIEDAD_PUBLICA
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
    End If
    
    'agregar propiedad al arbol
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_PROP_FRM & k, tvwChild, "FPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "FPRI_PROP" & NPROP
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_PROP_BAS & k, tvwChild, "BPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "BPRI_PROP" & NPROP
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_PROP_CLS & k, tvwChild, "CPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "CPRI_PROP" & NPROP
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_PROP_CTL & k, tvwChild, "KPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "KPRI_PROP" & NPROP
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_PROP_PAG & k, tvwChild, "PPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "PPRI_PROP" & NPROP
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_PROP_DSR & k, tvwChild, "DPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "DPRI_PROP" & NPROP
    End If
                    
    xproyecto.aArchivos(k).aRutinas(r).Linea = nLinea
    
    ReDim xproyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim xproyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    ReDim xproyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas(xproyecto)
    End If
    
    r = r + 1
    vr = 1
    nLinea = nLinea + 1
    prop = prop + 1
    NPROP = NPROP + 1
    
    If LineaPaso <> ")" Then
        Linea = LineaPaso
        Call ProcesarParametros(oTrv, xproyecto)
        If Right$(xproyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
            xproyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
        Else
            xproyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
            xproyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Propiedad)
        End If
    End If
    
End Sub

'analizar private sub
Private Sub AnalizaPrivateSub(oTrv As TreeView, xproyecto As eProyecto)
    
    Dim NombreSub As String
    Dim LineaX As String
    
    LineaX = NombreX(Linea)
    
    Procedimiento = NombreX(Linea)
    ReDim Preserve xproyecto.aArchivos(k).aRutinas(r)
    
    xproyecto.aArchivos(k).aRutinas(r).Nombre = Procedimiento
    Procedimiento = Mid$(Procedimiento, 13)
    
    xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Procedimiento, InStr(1, Procedimiento, "(") - 1)
    xproyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB
    xproyecto.aArchivos(k).aRutinas(r).Publica = False
    xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
    xproyecto.aArchivos(k).aRutinas(r).IsObjectSub = BuscaSubDeObjeto(oTrv, xproyecto, k, xproyecto.aArchivos(k).aRutinas(r).NombreRutina)
    xproyecto.aArchivos(k).aRutinas(r).Linea = nLinea
    
    xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
        
    ReDim xproyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim xproyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas(xproyecto)
    End If
    
    If Not bSub Then Call StartSubrutina(oTrv, xproyecto)
        
    NombreSub = xproyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    If Not bSubPri Then
        Call AgregaTipoDeSub(oTrv, xproyecto, False)
        bSubPri = True
    End If
    
    Call StartChildSub(oTrv, xproyecto, NombreSub, C_ICONO_PRIVATE_SUB, xproyecto.aArchivos(k).KeyNodeSub & "-SPRI" & PRISUB - 1)
                            
    ReDim xproyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
    nLinea = nLinea + 1
    r = r + 1
    s = s + 1
    spri = spri + 1
    PROC = PROC + 1
    vr = 1 'para contar variables rutinas
    
    'chequear si no viene la sub cortada
        
    Linea = Mid$(Linea, Len(LoadResString(C_PRIVATE_SUB)) + 1)
    Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)
    
    If Linea <> ")" Then
        Call ProcesarParametros(oTrv, xproyecto)
    End If
        
    Linea = LineaX
    
End Sub

'analizar public const
Private Sub AnalizaPublicConst(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreConstante As String
    
    Constante = NombreX(Linea)
    ReDim Preserve xproyecto.aArchivos(k).aConstantes(c)
    xproyecto.aArchivos(k).aConstantes(c).Nombre = Constante
    Constante = Mid$(Constante, 14)
    Constante = Left$(Constante, InStr(1, Constante, "=") - 2)
    
    If InStr(Constante, LoadResString(C_AS)) Then
        Constante = Left$(Constante, InStr(Constante, LoadResString(C_AS)) - 3)
    End If
        
    xproyecto.aArchivos(k).aConstantes(c).NombreVariable = Left$(Constante, InStr(1, Constante, "=") - 2)
    xproyecto.aArchivos(k).aConstantes(c).Publica = True
    xproyecto.aArchivos(k).aConstantes(c).Estado = NOCHEQUEADO
    xproyecto.aArchivos(k).aConstantes(c).Linea = nLinea
    xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
    
    If Not bCon Then Call StartConstantes(oTrv, xproyecto)
    
    NombreConstante = xproyecto.aArchivos(k).aConstantes(c).NombreVariable
    
    Call StartChildConstante(oTrv, xproyecto, NombreConstante)
    
    nLinea = nLinea + 1
    c = c + 1
    cpub = cpub + 1
    Cons = Cons + 1
                        
End Sub

'analiza public function
Private Sub AnalizaPublicFunction(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreFuncion As String
    
    Dim LineaX As String
    
    LineaX = NombreX(Linea)

    Funcion = NombreX(Linea)
    ReDim Preserve xproyecto.aArchivos(k).aRutinas(r)
    xproyecto.aArchivos(k).aRutinas(r).Nombre = Funcion
    
    If Right$(Funcion, 1) = ")" Then
        xproyecto.aArchivos(k).aRutinas(r).RegresaValor = False
    Else
        xproyecto.aArchivos(k).aRutinas(r).RegresaValor = True
        xproyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Funcion)
    End If
        
    Funcion = Mid$(Funcion, 17)
    xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Funcion, InStr(1, Funcion, "(") - 1)
    xproyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN
    xproyecto.aArchivos(k).aRutinas(r).Publica = True
    xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
    xproyecto.aArchivos(k).aRutinas(r).Linea = nLinea
    xproyecto.aArchivos(k).aRutinas(r).BasicStyle = BasicOldStyle(xproyecto.aArchivos(k).aRutinas(r).NombreRutina)
    xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
    
    ReDim xproyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim xproyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas(xproyecto)
    End If
    
    If Not bFun Then Call StartFuncion(oTrv, xproyecto)
    
    NombreFuncion = xproyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    If Not bFunPub Then
        Call AgregaTipoDeFuncion(oTrv, xproyecto, True)
        bFunPub = True
    End If
    
    Call StartChildFuncion(oTrv, xproyecto, NombreFuncion, C_ICONO_PUBLIC_FUNCION, xproyecto.aArchivos(k).KeyNodeFun & "-FPUB" & PUBFUN - 1)
    
    ReDim xproyecto.aArchivos(k).aRutinas(r).Aparams(0)
            
    nLinea = nLinea + 1
    r = r + 1
    f = f + 1
    fpub = fpub + 1
    Func = Func + 1
    vr = 1 'para contar variables rutinas
    
    'chequear si no viene la fun cortada
    Linea = Mid$(Linea, Len(LoadResString(C_PUBLIC_FUNCTION)) + 1)
    Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)
    
    If Linea <> ")" Then
        Call ProcesarParametros(oTrv, xproyecto)
        If Right$(xproyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
            xproyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
        Else
            xproyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
            xproyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Funcion)
        End If
    End If
        
    Linea = LineaX
    
End Sub
'analiza public sub
Private Sub AnalizaPublicSub(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreSub As String
    Dim LineaX As String
    
    LineaX = NombreX(Linea)
        
    Procedimiento = NombreX(Linea)
    ReDim Preserve xproyecto.aArchivos(k).aRutinas(r)
    xproyecto.aArchivos(k).aRutinas(r).Nombre = Procedimiento
    Procedimiento = Mid$(Procedimiento, 12)
    xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Procedimiento, InStr(1, Procedimiento, "(") - 1)
    xproyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB
    xproyecto.aArchivos(k).aRutinas(r).Publica = True
    xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
    xproyecto.aArchivos(k).aRutinas(r).Linea = nLinea
    xproyecto.aArchivos(k).aRutinas(r).IsObjectSub = BuscaSubDeObjeto(oTrv, xproyecto, k, xproyecto.aArchivos(k).aRutinas(r).NombreRutina)
    
    xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
    
    ReDim xproyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim xproyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas(xproyecto)
    End If
    
    If Not bSub Then Call StartSubrutina(oTrv, xproyecto)
    
    NombreSub = xproyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    If Not bSubPub Then
        Call AgregaTipoDeSub(oTrv, xproyecto, True)
        bSubPub = True
    End If
    
    Call StartChildSub(oTrv, xproyecto, NombreSub, C_ICONO_PUBLIC_SUB, xproyecto.aArchivos(k).KeyNodeSub & "-SPUB" & PUBSUB - 1)
    
    ReDim xproyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
    nLinea = nLinea + 1
    r = r + 1
    s = s + 1
    spub = spub + 1
    PROC = PROC + 1
    vr = 1 'para contar variables rutinas
    
    'chequear si no viene la sub cortada
    Linea = Mid$(Linea, Len(LoadResString(C_PUBLIC_SUB)) + 1)
    Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)
    
    If Linea <> ")" Then
        Call ProcesarParametros(oTrv, xproyecto)
    End If
                            
    Linea = LineaX
    
End Sub

'analiza sub
Private Sub AnalizaSub(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreSub As String
    Dim LineaX As String
    
    LineaX = NombreX(Linea)
    
    If Left$(Linea, 3) = Trim$(LoadResString(C_SUB)) Or Left$(Linea, 10) = Trim$(LoadResString(C_FRIEND_SUB)) Then
                
        Procedimiento = Linea
        ReDim Preserve xproyecto.aArchivos(k).aRutinas(r)
        xproyecto.aArchivos(k).aRutinas(r).Nombre = Procedimiento
        
        If Left$(Linea, 3) = Trim$(LoadResString(C_SUB)) Then
            Procedimiento = Mid$(Procedimiento, 5)
        ElseIf Left$(Linea, 10) = Trim$(LoadResString(C_FRIEND_SUB)) Then
            Procedimiento = Mid$(Procedimiento, 12)
        End If
        
        xproyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Procedimiento, InStr(1, Procedimiento, "(") - 1)
        xproyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB
        xproyecto.aArchivos(k).aRutinas(r).Publica = True
        xproyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).aRutinas(r).IsObjectSub = BuscaSubDeObjeto(oTrv, xproyecto, k, xproyecto.aArchivos(k).aRutinas(r).NombreRutina)
        xproyecto.aArchivos(k).aRutinas(r).Linea = nLinea
        
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
        
        ReDim xproyecto.aArchivos(k).aRutinas(r).aVariables(0)
        ReDim xproyecto.aArchivos(k).aRutinas(r).aRVariables(0)
        
        If Not StartRutinas Then
            StartRutinas = True
            Call StartRastreoRutinas(xproyecto)
        End If
        
        If Not bSub Then Call StartSubrutina(oTrv, xproyecto)
    
        If Not bSubPub Then
            Call AgregaTipoDeSub(oTrv, xproyecto, True)
            bSubPub = True
        End If
    
        NombreSub = xproyecto.aArchivos(k).aRutinas(r).NombreRutina
        
        Call StartChildSub(oTrv, xproyecto, NombreSub, C_ICONO_PUBLIC_SUB, xproyecto.aArchivos(k).KeyNodeSub & "-SPUB" & PUBSUB - 1)
    
        ReDim xproyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
        nLinea = nLinea + 1
        r = r + 1
        s = s + 1
        spub = spub + 1
        PROC = PROC + 1
        vr = 1 'para contar variables rutinas
        
        'chequear si no viene la sub cortada
        Linea = Mid$(Linea, Len(LoadResString(C_SUB)) + 1)
        Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)
        
        If Linea <> ")" Then
            Call ProcesarParametros(oTrv, xproyecto)
        End If
    ElseIf InStr(Linea, LoadResString(C_DECLARE)) <> 0 Then    'API
        Call AnalizaApi(oTrv, xproyecto)
    End If
                        
    Linea = LineaX
    
End Sub

'analizar tipos
Private Sub AnalizaType(oTrv As TreeView, xproyecto As eProyecto)

    Dim NombreTipo As String
    
    If Left$(Linea, 5) = LoadResString(C_TYPE) Then
        
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
        End If
                            
        StartTypes = True
        
        Tipo = Linea
        ReDim Preserve xproyecto.aArchivos(k).aTipos(t)
        
        ReDim xproyecto.aArchivos(k).aTipos(t).aElementos(0)
        
        xproyecto.aArchivos(k).aTipos(t).Nombre = Tipo
        Tipo = Mid$(Tipo, 6)
        xproyecto.aArchivos(k).aTipos(t).NombreVariable = Tipo
        xproyecto.aArchivos(k).aTipos(t).Publica = False
        xproyecto.aArchivos(k).aTipos(t).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).aTipos(t).Linea = nLinea
        xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
        
        If Not bTipo Then Call StartTipos(oTrv, xproyecto)
        
        NombreTipo = xproyecto.aArchivos(k).aTipos(t).NombreVariable
        
        Call StartChildTipos(oTrv, xproyecto, NombreTipo)
        
        t = t + 1
        TYPO = TYPO + 1
        tpub = tpub + 1
        nLinea = nLinea + 1
    ElseIf Left$(Linea, 11) = LoadResString(C_PUBLIC_TYPE) Then
        
        StartTypes = True
        
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
        End If
        
        Tipo = Linea
        ReDim Preserve xproyecto.aArchivos(k).aTipos(t)
        
        ReDim xproyecto.aArchivos(k).aTipos(t).aElementos(0)
        
        xproyecto.aArchivos(k).aTipos(t).Nombre = NombreX(Tipo)
        Tipo = Mid$(Tipo, 13)
        xproyecto.aArchivos(k).aTipos(t).NombreVariable = NombreX(Tipo)
        xproyecto.aArchivos(k).aTipos(t).Publica = True
        xproyecto.aArchivos(k).aTipos(t).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).aTipos(t).Linea = nLinea
        xproyecto.aArchivos(k).MiembrosPublicos = xproyecto.aArchivos(k).MiembrosPublicos + 1
        
        If Not bTipo Then Call StartTipos(oTrv, xproyecto)
        
        NombreTipo = xproyecto.aArchivos(k).aTipos(t).NombreVariable
        
        Call StartChildTipos(oTrv, xproyecto, NombreTipo)
        
        t = t + 1
        TYPO = TYPO + 1
        tpub = tpub + 1
        nLinea = nLinea + 1
    ElseIf Left$(Linea, 12) = LoadResString(C_PRIVATE_TYPE) Then
        
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
        End If
        
        'para comenzar a guardar los elementos del tipo
        If Not StartTypes Then
            StartTypes = True
        End If
        
        Tipo = Linea
        ReDim Preserve xproyecto.aArchivos(k).aTipos(t)
        
        ReDim xproyecto.aArchivos(k).aTipos(t).aElementos(0)
        
        xproyecto.aArchivos(k).aTipos(t).Nombre = NombreX(Tipo)
        Tipo = Mid$(Tipo, 14)
        xproyecto.aArchivos(k).aTipos(t).NombreVariable = NombreX(Tipo)
        xproyecto.aArchivos(k).aTipos(t).Publica = False
        xproyecto.aArchivos(k).aTipos(t).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).aTipos(t).Linea = nLinea
        xproyecto.aArchivos(k).MiembrosPrivados = xproyecto.aArchivos(k).MiembrosPrivados + 1
        
        If Not bTipo Then Call StartTipos(oTrv, xproyecto)
        
        NombreTipo = xproyecto.aArchivos(k).aTipos(t).NombreVariable
        
        Call StartChildTipos(oTrv, xproyecto, NombreTipo)
        
        t = t + 1
        TYPO = TYPO + 1
        tpri = tpri + 1
        nLinea = nLinea + 1
    End If
                        
End Sub

'comprueba si la variable esta declarada al viejo estilo
'de basic
Private Function BasicOldStyle(ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If Right$(Variable, 1) = "$" Then
        ret = True
    ElseIf Right$(Variable, 1) = "!" Then
        ret = True
    ElseIf Right$(Variable, 1) = "#" Then
        ret = True
    ElseIf Right$(Variable, 1) = "@" Then
        ret = True
    ElseIf Right$(Variable, 1) = "&" Then
        ret = True
    ElseIf Right$(Variable, 1) = "%" Then
        ret = True
    End If
    
    BasicOldStyle = ret
    
End Function

'determina si la sub es de un objeto
Private Function BuscaSubDeObjeto(oTrv As TreeView, xproyecto As eProyecto, ByVal k As Integer, ByVal Subrutina As String) As Boolean

    Dim ret As Boolean
    Dim ca As Integer
    Dim sMenu As String
    Dim j As Integer
    
    If Left$(LCase$(Subrutina), 4) = LCase$("Form") Or Left$(LCase$(Subrutina), 7) = LCase$("MDIForm") Then
        ret = True
        GoTo Salir
    ElseIf Left$(LCase$(Subrutina), 11) = LCase$("UserControl") Then
        ret = True
        GoTo Salir
    End If
        
    ret = False
    
    If InStr(Subrutina, "_") = 0 Then
        GoTo Salir
    End If
    
    'sacar el evento de la rutina. si es que existe
    For j = Len(Subrutina) To 1 Step -1
        If Mid$(Subrutina, j, 1) = "_" Then
            sMenu = UCase$(Trim$(Left$(Subrutina, j - 1)))
            Exit For
        End If
    Next j
            
    'ciclar x los controles
    For ca = 1 To UBound(xproyecto.aArchivos(k).aControles())
        If UCase$(Trim$(xproyecto.aArchivos(k).aControles(ca).Nombre)) = sMenu Then
            ret = True
            Exit For
        End If
    Next ca
    
Salir:
    BuscaSubDeObjeto = ret
    
End Function

'cargar archivos requeridos por el xproyecto dlls, ocxs, res
Private Sub CargaArchivosxproyecto(oTrv As TreeView, xproyecto As eProyecto)

    Dim k As Integer
    
    Dim bReferencias As Boolean
    Dim bOcxs As Boolean
    Dim bRes As Boolean
    Dim bPags As Boolean
    Dim bForm As Boolean
    Dim bModule As Boolean
    Dim bControl As Boolean
    Dim bClase As Boolean
    Dim bDocRel As Boolean
    Dim bDesigner As Boolean
    
    Call HelpCarga(LoadResString(C_PROPIEDADES_PROYECTO))
    
    For k = 1 To UBound(xproyecto.aDepencias)
        If xproyecto.aDepencias(k).Tipo = TIPO_DLL Then
            If Not bReferencias Then
                Call oTrv.Nodes.Add("PRO", tvwChild, "REFDLL", "Referencias", C_ICONO_CLOSE).EnsureVisible
                bReferencias = True
            End If
            
            'archivo .dll
            Call oTrv.Nodes.Add("REFDLL", tvwChild, xproyecto.aDepencias(k).KeyNode, xproyecto.aDepencias(k).ContainingFile, C_ICONO_REFERENCIAS, C_ICONO_REFERENCIAS)
            
            'informacion de esta
            Call oTrv.Nodes.Add(xproyecto.aDepencias(k).KeyNode, tvwChild, , xproyecto.aDepencias(k).HelpString, C_ICONO_ARCHIVO_REF, C_ICONO_ARCHIVO_REF)
            Call oTrv.Nodes.Add(xproyecto.aDepencias(k).KeyNode, tvwChild, , xproyecto.aDepencias(k).GUID, C_ICONO_ARCHIVO_REF, C_ICONO_ARCHIVO_REF)
        ElseIf xproyecto.aDepencias(k).Tipo = TIPO_OCX Then
            If Not bOcxs Then
                Call oTrv.Nodes.Add("PRO", tvwChild, "REFOCX", "Componentes", C_ICONO_CLOSE).EnsureVisible
                bOcxs = True
            End If
            'archivo .ocx
            Call oTrv.Nodes.Add("REFOCX", tvwChild, xproyecto.aDepencias(k).KeyNode, xproyecto.aDepencias(k).ContainingFile, C_ICONO_OCX, C_ICONO_OCX)
            
            'informacion de esta
            Call oTrv.Nodes.Add(xproyecto.aDepencias(k).KeyNode, tvwChild, , xproyecto.aDepencias(k).HelpString, C_ICONO_CONTROL, C_ICONO_CONTROL)
            Call oTrv.Nodes.Add(xproyecto.aDepencias(k).KeyNode, tvwChild, , xproyecto.aDepencias(k).GUID, C_ICONO_CONTROL, C_ICONO_CONTROL)
            
        ElseIf xproyecto.aDepencias(k).Tipo = TIPO_RES Then
            If Not bRes Then
                Call oTrv.Nodes.Add("PRO", tvwChild, "REFRES", "Recursos", C_ICONO_CLOSE).EnsureVisible
                bRes = True
            End If
            Call oTrv.Nodes.Add("REFRES", tvwChild, xproyecto.aDepencias(k).KeyNode, xproyecto.aDepencias(k).Archivo, C_ICONO_RECURSO, C_ICONO_RECURSO)
        End If
    Next k
        
    'cargar archivos del xproyecto
    For k = 1 To UBound(xproyecto.aArchivos)
        If xproyecto.aArchivos(k).Explorar = True Then
            If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                If Not bForm Then
                    Call oTrv.Nodes.Add("PRO", tvwChild, "FRM", "Formularios", C_ICONO_CLOSE).EnsureVisible
                    bForm = True
                End If
                Call oTrv.Nodes.Add("FRM", tvwChild, xproyecto.aArchivos(k).KeyNodeFrm, xproyecto.aArchivos(k).Nombre, C_ICONO_FORM, C_ICONO_FORM)
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                If Not bModule Then
                    Call oTrv.Nodes.Add("PRO", tvwChild, "BAS", "Módulos", C_ICONO_CLOSE).EnsureVisible
                    bModule = True
                End If
                Call oTrv.Nodes.Add("BAS", tvwChild, xproyecto.aArchivos(k).KeyNodeBas, xproyecto.aArchivos(k).Nombre, C_ICONO_BAS, C_ICONO_BAS)
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                If Not bControl Then
                    Call oTrv.Nodes.Add("PRO", tvwChild, "CTL", "Controles de Usuario", C_ICONO_CLOSE).EnsureVisible
                    bControl = True
                End If
                Call oTrv.Nodes.Add("CTL", tvwChild, xproyecto.aArchivos(k).KeyNodeKtl, xproyecto.aArchivos(k).Nombre, C_ICONO_CONTROL, C_ICONO_CONTROL)
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                If Not bClase Then
                    Call oTrv.Nodes.Add("PRO", tvwChild, "CLS", "Módulos de Clase", C_ICONO_CLOSE).EnsureVisible
                    bClase = True
                End If
                Call oTrv.Nodes.Add("CLS", tvwChild, xproyecto.aArchivos(k).KeyNodeCls, xproyecto.aArchivos(k).Nombre, C_ICONO_CLS, C_ICONO_CLS)
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                If Not bPags Then
                    Call oTrv.Nodes.Add("PRO", tvwChild, "PAG", "Páginas de Propiedades", C_ICONO_CLOSE).EnsureVisible
                    bPags = True
                End If
                Call oTrv.Nodes.Add("PAG", tvwChild, xproyecto.aArchivos(k).KeyNodePag, xproyecto.aArchivos(k).Nombre, C_ICONO_PAGINA, C_ICONO_PAGINA)
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_REL Then
                If Not bDocRel Then
                    Call oTrv.Nodes.Add("PRO", tvwChild, "REL", "Documentos Relacionados", C_ICONO_CLOSE).EnsureVisible
                    bDocRel = True
                End If
                Call oTrv.Nodes.Add("REL", tvwChild, xproyecto.aArchivos(k).KeyNodeRel, xproyecto.aArchivos(k).Nombre, C_ICONO_DOCREL, C_ICONO_DOCREL)
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                If Not bDesigner Then
                    Call oTrv.Nodes.Add("PRO", tvwChild, C_KEY_DSR, "Diseñadores", C_ICONO_CLOSE).EnsureVisible
                    bDesigner = True
                End If
                Call oTrv.Nodes.Add(C_KEY_DSR, tvwChild, xproyecto.aArchivos(k).KeyNodeDsr, xproyecto.aArchivos(k).Nombre, C_ICONO_DESIGNER, C_ICONO_DESIGNER)
            End If
            
        End If
    Next k
    
End Sub

Public Function CargaProyecto(ByVal Archivo As String, oTrv As TreeView, _
                              xproyecto As eProyecto, xTotalesProyecto As eTotalesProyecto) As Boolean

    'On Local Error GoTo ErrorCargaxproyecto
    
    Dim ret As Boolean
    Dim Tipoxproyecto As Long
    Dim Icono As Long
        
    Dim Linea As String
        
    Dim f As Integer
    Dim M As Integer
    Dim c As Integer
    Dim k As Integer
    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
    Dim d As Integer
        
    Dim Formulario As String
    Dim Modulo As String
    Dim ControlUsuario As String
    Dim Clase As String
    Dim Referencia As String
    Dim RefRes As String
    Dim PagPropiedades As String
    Dim DoctosRelacionados As String
    Dim Disenador As String
    Dim nFreeFile As Long
    
    Dim sSystem As String
    
    ret = True
    
    oTrv.Nodes.Clear
            
    Archivo = StripNulls(Archivo)
    
    Pathxproyecto = PathArchivo(Archivo)
    
    f = 1 'frm
    M = 1 'bas
    c = 1 'cls
    k = 1 'ctl
    d = 1 '->dependencias ...
            
    nFreeFile = FreeFile
        
    Call HelpCarga(LoadResString(C_LEYENDO_ARCHIVOS))
        
    glbSeComparo = False
    
    'determinar el tipo de xproyecto
    If Not DeterminaTipoDeproyecto(oTrv, xproyecto, Archivo) Then
        ret = False
        GoTo SalirCargaxproyecto
    End If
    
    xproyecto.PathFisico = Archivo
    xproyecto.FILETIME = VBGetFileTime(Archivo)
    
    If oTrv.Name = "treeProyectoO" Then
        frmMain.lblOrigen.Caption = "Origen : <" & xproyecto.FILETIME & ">"
    Else
        frmMain.lblDestino.Caption = "Destino : <" & xproyecto.FILETIME & ">"
    End If
    
    ReDim xproyecto.aArchivos(0)
    ReDim xproyecto.aDepencias(0)
    
    nFreeFile = FreeFile
    
    'limpiar acumuladores generales
    Call LimpiarTotales(xTotalesProyecto)
    
    REF_DLL = 1
    REF_OCX = 1
    REF_RES = 1
            
    'determinar los diferentes archivos que componen el xproyecto
    Open Archivo For Input Shared As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, Linea
            If InStr(Linea, "Form=") Then           'FORMULARIOS
                If InStr(Linea, "IconForm=") = 0 Then
                    Formulario = Mid$(Linea, InStr(Linea, "=") + 1)
                    
                    Call AgregaArchivoDexproyecto(oTrv, xproyecto, k, Formulario, TIPO_ARCHIVO_FRM, C_KEY_FRM)
                End If
            ElseIf InStr(Linea, "Module=") Then     'MODULOS
                Modulo = Mid$(Linea, InStr(Linea, "=") + 1)
                Modulo = Trim$(Mid$(Modulo, InStr(Modulo, ";") + 1))
                
                Call AgregaArchivoDexproyecto(oTrv, xproyecto, k, Modulo, TIPO_ARCHIVO_BAS, C_KEY_BAS)
            ElseIf InStr(Linea, "UserControl=") Then    'CONTROLES
                ControlUsuario = Mid$(Linea, InStr(Linea, "=") + 1)
                Call AgregaArchivoDexproyecto(oTrv, xproyecto, k, ControlUsuario, TIPO_ARCHIVO_OCX, C_KEY_CTL)
            ElseIf InStr(Linea, "Class=") Then          'MODULOS DE CLASE
                Clase = Mid$(Linea, InStr(Linea, "=") + 1)
                Clase = Trim$(Mid$(Clase, InStr(Clase, ";") + 1))
                                                
                Call AgregaArchivoDexproyecto(oTrv, xproyecto, k, Clase, TIPO_ARCHIVO_CLS, C_KEY_CLS)
            ElseIf InStr(Linea, "Reference=") Then      'REFERENCIAS
                Call AgregaReferencias(oTrv, xproyecto, d, Linea)
            ElseIf InStr(Linea, "Object=") Then         'CONTROLES
                If Left$(Linea, 6) = "Object" Then
                    Call AgregaComponentes(oTrv, xproyecto, d, Linea)
                End If
            ElseIf InStr(Linea, "ResFile32=") Then
                RefRes = Trim$(Mid$(Linea, InStr(Linea, """") + 1))
                RefRes = Left$(RefRes, Len(RefRes) - 1)
                
                ReDim Preserve xproyecto.aDepencias(d)
                
                'CHEQUEAR \
                If PathArchivo(RefRes) = "" Then
                    xproyecto.aDepencias(d).Archivo = Pathxproyecto & RefRes
                Else
                    xproyecto.aDepencias(d).Archivo = PathArchivo(RefRes)
                End If
                
                xproyecto.aDepencias(d).Tipo = TIPO_RES
                xproyecto.aDepencias(d).KeyNode = "REFRES" & REF_RES
                xproyecto.aDepencias(d).FileSize = VBGetFileSize(xproyecto.aDepencias(d).Archivo)
                xproyecto.aDepencias(d).FILETIME = VBGetFileTime(xproyecto.aDepencias(d).Archivo)
                REF_RES = REF_RES + 1
                d = d + 1
            ElseIf InStr(Linea, "PropertyPage=") Then   'Pagina de propiedades
                PagPropiedades = Mid$(Linea, InStr(Linea, "=") + 1)
                
                Call AgregaArchivoDexproyecto(oTrv, xproyecto, k, PagPropiedades, TIPO_ARCHIVO_PAG, C_KEY_PAG)
            ElseIf InStr(Linea, "Designer=") Then   'diseñadores
                Disenador = Mid$(Linea, InStr(Linea, "=") + 1)
                
                Call AgregaArchivoDexproyecto(oTrv, xproyecto, k, Disenador, TIPO_ARCHIVO_DSR, C_KEY_DSR)
            ElseIf InStr(Linea, "RelatedDoc=") Then   'Documentos Relacionados
                DoctosRelacionados = Mid$(Linea, InStr(Linea, "=") + 1)
                
                Call AgregaArchivoDexproyecto(oTrv, xproyecto, k, DoctosRelacionados, TIPO_ARCHIVO_REL, C_KEY_REL)
            
            ElseIf InStr(Linea, "Startup") Then
                If Left$(Linea, 8) = "Startup=" Then
                    xproyecto.Startup = Mid$(Linea, InStr(Linea, "=") + 2)
                    xproyecto.Startup = Left$(xproyecto.Startup, Len(xproyecto.Startup) - 1)
                End If
            ElseIf Right$(Linea, 3) = "FRM" Then 'para versiones anteriores de VB3
                Formulario = Linea
                Call AgregaArchivoDexproyecto(oTrv, xproyecto, k, Formulario, TIPO_ARCHIVO_FRM, C_KEY_FRM)
            ElseIf Right$(Linea, 3) = "BAS" Then 'para versiones anteriores de VB3
                Modulo = Linea
                Call AgregaArchivoDexproyecto(oTrv, xproyecto, k, Modulo, TIPO_ARCHIVO_BAS, C_KEY_BAS)
            ElseIf Right$(Linea, 3) = "VBX" Then 'para versiones anteriores de VB3
                Call AgregaReferencias(oTrv, xproyecto, d, Linea)
            ElseIf Left$(Linea, 10) = "ExeName32=" Then
                Linea = Mid$(Linea, 12)
                If InStr(Linea, "\") Then
                    xproyecto.ExeName = VBArchivoSinPath(Linea)
                Else
                    xproyecto.ExeName = Left$(Linea, Len(Linea) - 1)
                End If
            End If
        Loop
    Close #nFreeFile
        
    If oTrv.Name = "treeProyectoO" Then
        frmSelExplorar.Origen = True
    Else
        frmSelExplorar.Origen = False
    End If
    
    frmSelExplorar.Show vbModal
    
    If glbSelArchivos Then
        Call ShowProgress(True)
        Call Hourglass(frmMain.hWnd, True)
        oTrv.Nodes.Add(, , "PRO", xproyecto.Nombre & " (" & xproyecto.Archivo & ")", xproyecto.Icono).EnsureVisible
        Call CargaArchivosxproyecto(oTrv, xproyecto)
        Call AnalizaArchivosDelproyecto(oTrv, xproyecto, xTotalesProyecto)
        Call DeterminaEventosControles(oTrv, xproyecto)
        Call SeteaContadoresAnalisis(oTrv, xproyecto)
        
        MsgBox xproyecto.Nombre & " " & LoadResString(C_EXITO_CARGA), vbInformation
        
        Call Hourglass(frmMain.hWnd, False)
    Else
        ret = False
    End If
        
    GoTo SalirCargaxproyecto
    
ErrorCargaxproyecto:
    ret = False
    MsgBox "Cargaxproyecto : " & Err & " " & Error$, vbCritical
    Resume SalirCargaxproyecto
    
SalirCargaxproyecto:
    Set cRegistro = Nothing
    Set cTLI = Nothing
    
    Call ShowProgress(False)
    CargaProyecto = ret
    Call HelpCarga(LoadResString(C_LISTO))
    frmMain.stbMain.Panels(2).Text = ""
    frmMain.stbMain.Panels(4).Text = ""
    Err = 0
    
End Function

'comprueba si se comienza el desglose de funcion sub
Private Function ChequeaDesgloseSubFun(ByVal LineaX As String) As Boolean

    Dim ret As Boolean
    
    ret = False
        
    If InStr(LineaX, LoadResString(C_PRIVATE_SUB)) Then
        If Left$(LineaX, 12) = LoadResString(C_PRIVATE_SUB) Then     'PRIVATE SUB
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_PUBLIC_SUB)) Then
        If Left$(LineaX, 11) = LoadResString(C_PUBLIC_SUB) Then      'PUBLIC SUB
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_FRIEND_SUB)) Then
        If Left$(LineaX, 11) = LoadResString(C_FRIEND_SUB) Then        'FRIEND SUB
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_SUB)) Then
        If Left$(LineaX, 4) = LoadResString(C_SUB) Then             'SUB
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_PRIVATE_FUNCTION)) Then
        If Left$(LineaX, 17) = LoadResString(C_PRIVATE_FUNCTION) Then 'PRIVATE FUNCTION
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_PUBLIC_FUNCTION)) Then
        If Left$(LineaX, 16) = LoadResString(C_PUBLIC_FUNCTION) Then 'PUBLIC FUNCTION
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_FUNCTION)) Then
        If Left$(LineaX, 9) = LoadResString(C_FUNCTION) Then        'FUNCTION
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_FRIEND_FUNCTION)) Then
        If Left$(LineaX, 16) = LoadResString(C_FRIEND_FUNCTION) Then   'FRIEND FUNCTION
            ret = True
            EndGeneral = True
        End If
    Else
        ret = False
    End If
                                    
    ChequeaDesgloseSubFun = ret
    
End Function
'comprueba la continuacion de linea o el espacio en blanco
Private Sub ChequeaLineaDeRutina(xproyecto As eProyecto, ByVal FlagLinea As Boolean)

    If StartRutinas Then
        Call GrabaLineaDeRutina(xproyecto)
        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
        
        If LineaOrigen = "" Then           'espacios en blancos
            xproyecto.aArchivos(k).NumeroDeLineasEnBlanco = xproyecto.aArchivos(k).NumeroDeLineasEnBlanco + 1
            
            xproyecto.aArchivos(k).aRutinas(r - 1).NumeroDeBlancos = _
            xproyecto.aArchivos(k).aRutinas(r - 1).NumeroDeBlancos + 1
        ElseIf Left$(LineaOrigen, 1) = "'" Then  'comentarios
            xproyecto.aArchivos(k).NumeroDeLineasComentario = xproyecto.aArchivos(k).NumeroDeLineasComentario + 1
            
            xproyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios = _
            xproyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios + 1
        End If
                        
        'total de lineas de la rutina
        xproyecto.aArchivos(k).aRutinas(r - 1).TotalLineas = xproyecto.aArchivos(k).aRutinas(r - 1).TotalLineas + 1
    ElseIf StartGeneral Then
        If Not EndGeneral Then
            If Not FlagLinea Then
                ReDim Preserve xproyecto.aArchivos(k).aGeneral(ge)
                xproyecto.aArchivos(k).aGeneral(ge).Codigo = LineaOrigen
                xproyecto.aArchivos(k).aGeneral(ge).Linea = nLinea
                ge = ge + 1
                xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
            Else
                If Not ChequeaDesgloseSubFun(LineaOrigen) Then
                    ReDim Preserve xproyecto.aArchivos(k).aGeneral(ge)
                    xproyecto.aArchivos(k).aGeneral(ge).Codigo = LineaOrigen
                    xproyecto.aArchivos(k).aGeneral(ge).Linea = nLinea
                    ge = ge + 1
                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                End If
            End If
        End If
    End If
                                        
End Sub
Public Function ClearEnterInString(ByVal sText As String) As String

    Dim k As Integer
    
    Dim ret As String
    
    For k = 1 To Len(sText)
        If Chr$(Asc(Mid$(sText, k, 1))) <> Chr$(13) And Chr$(Asc(Mid$(sText, k, 1))) <> Chr$(10) Then
            ret = ret & Mid$(sText, k, 1)
        Else
            ret = ret & " "
        End If
    Next k
    
    ClearEnterInString = ret
    
End Function
'Cargar datos de modulo
Private Sub AnalizaArchivosDelproyecto(oTrv As TreeView, xproyecto As eProyecto, xTotalesProyecto As eTotalesProyecto)
        
    Dim j As Integer
    Dim sNombre As String
    Dim FlagLinea As Boolean
    Dim fFirstRutina As Boolean
    
    Call InicializarVariables(xproyecto)
    
    ValidateRect oTrv.hWnd, 0&
    
    For k = 1 To UBound(xproyecto.aArchivos)
        If xproyecto.aArchivos(k).Explorar Then
            DoEvents
            
            Call InicializarVariablesArchivos(xproyecto)
                                    
            'abrir archivo en proceso
            If VBOpenFile(xproyecto.aArchivos(k).PathFisico) Then
            
                ReDim Arr_Paso(0)
                fFirstRutina = False
                Open xproyecto.aArchivos(k).PathFisico For Input Shared As #nFreeFile
                    Do While Not EOF(nFreeFile)
                        Line Input #nFreeFile, Linea
                        
                        LineaOrigen = Linea
                        Linea = Trim$(Linea)

                        'linea en blanco
                        If Linea <> "" Then
                            'continuación de linea ?
                            If Right$(Linea, 1) = "_" Then
                                If EndGeneral Then
                                    If r > 1 Or fFirstRutina Then
                                        ReDim Preserve Arr_Paso(UBound(Arr_Paso) + 1)
                                        Arr_Paso(UBound(Arr_Paso)) = LineaOrigen
                                    Else
                                        ReDim Preserve Arr_Paso(UBound(Arr_Paso) + 1)
                                        Arr_Paso(UBound(Arr_Paso)) = LineaPaso
                                        
                                        ReDim Preserve Arr_Paso(UBound(Arr_Paso) + 1)
                                        Arr_Paso(UBound(Arr_Paso)) = LineaOrigen
                                        
                                        fFirstRutina = True
                                    End If
                                End If
                                
                                LineaPaso = LineaPaso & Left$(Linea, Len(Linea) - 1)
                                Linea = ""
                                FlagLinea = True
                            ElseIf LineaPaso <> "" Then
                                                                                                
                                If EndGeneral Then
                                    ReDim Preserve Arr_Paso(UBound(Arr_Paso) + 1)
                                    Arr_Paso(UBound(Arr_Paso)) = LineaOrigen
                                End If
                                
                                LineaPaso = LineaPaso & Linea
                                Linea = LineaPaso
                                LineaPaso = vbNullString
                                FlagLinea = False
                            End If
                            
                            'analizar linea ?
                            If Linea <> "" Then
                                ValidateRect oTrv.hWnd, 0&
                                
                                If InStr(Linea, LoadResString(C_OPTION_EXPLICIT)) Then       'OPTION EXPLICIT
                                    xproyecto.aArchivos(k).OptionExplicit = True
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                        nLinea = 1
                                    End If
                                ElseIf Left$(Linea, 1) = "'" Then   'COMENTARIO
                                    xproyecto.aArchivos(k).NumeroDeLineasComentario = xproyecto.aArchivos(k).NumeroDeLineasComentario + 1
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'agregar contador a rutinas
                                    If StartRutinas Then
                                        xproyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios = _
                                        xproyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios + 1
                                    End If
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                        nLinea = 1
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_DECLARE_FUNCTION)) Or InStr(Linea, LoadResString(C_PUBLIC_DECLARE_FUNCTION)) Or InStr(Linea, LoadResString(C_DECLARE_FUNCTION)) Then
                                    If Left$(Linea, 23) = LoadResString(C_PUBLIC_DECLARE_FUNCTION) Then    'PUBLIC DECLARE FUNCTION
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                            nLinea = 1
                                        End If
                                        Call AnalizaApi(oTrv, xproyecto)
                                    ElseIf Left$(Linea, 24) = LoadResString(C_PRIVATE_DECLARE_FUNCTION) Then   'PRIVATE DECLARE FUNCTION
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                            nLinea = 1
                                        End If
                                        Call AnalizaApi(oTrv, xproyecto)
                                    ElseIf Left$(Linea, 16) = LoadResString(C_DECLARE_FUNCTION) Then   'DECLARE FUNCTION
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                            nLinea = 1
                                        End If
                                        Call AnalizaApi(oTrv, xproyecto)
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_DECLARE_SUB)) Or InStr(Linea, LoadResString(C_PUBLIC_DECLARE_SUB)) Or InStr(Linea, LoadResString(C_DECLARE_SUB)) Then
                                    If Left$(Linea, 18) = LoadResString(C_PUBLIC_DECLARE_SUB) Then     'PUBLIC DECLARE SUB
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                            nLinea = 1
                                        End If
                                        
                                        Call AnalizaApi(oTrv, xproyecto)
                                    ElseIf Left$(Linea, 19) = LoadResString(C_PRIVATE_DECLARE_SUB) Then    'PRIVATE DECLARE SUB
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                            nLinea = 1
                                        End If
                                        
                                        Call AnalizaApi(oTrv, xproyecto)
                                    ElseIf Left$(Linea, 11) = LoadResString(C_DECLARE_SUB) Then    'DECLARE SUB
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                            nLinea = 1
                                        End If
                                        
                                        Call AnalizaApi(oTrv, xproyecto)
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_SUB)) Then
                                    If Left$(Linea, 12) = LoadResString(C_PRIVATE_SUB) Then     'PRIVATE SUB
                                        EndGeneral = True
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        nLinea = 1
                                        Call AnalizaPrivateSub(oTrv, xproyecto)
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_END_SUB)) Then             'END SUB
                                    If Left$(Linea, 7) = LoadResString(C_END_SUB) Then
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        Call FinalizarSub(xproyecto)
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PUBLIC_SUB)) Then
                                    If Left$(Linea, 11) = LoadResString(C_PUBLIC_SUB) Then      'PUBLIC SUB
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        EndGeneral = True
                                        nLinea = 1
                                        Call AnalizaPublicSub(oTrv, xproyecto)
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_FRIEND_SUB)) Then
                                    If Left$(Linea, 11) = LoadResString(C_FRIEND_SUB) Then        'FRIEND SUB
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        EndGeneral = True
                                        nLinea = 1
                                        Call AnalizaSub(oTrv, xproyecto)
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_SUB)) Then
                                    If Left$(Linea, 4) = LoadResString(C_SUB) Then             'SUB
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        EndGeneral = True
                                        nLinea = 1
                                        Call AnalizaSub(oTrv, xproyecto)
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_FUNCTION)) And Left$(Linea, 17) = LoadResString(C_PRIVATE_FUNCTION) Then 'PRIVATE FUNCTION
                                    EndGeneral = True
                                    nLinea = 1
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaPrivateFunction(oTrv, xproyecto)
                                ElseIf InStr(Linea, LoadResString(C_PUBLIC_FUNCTION)) And Left$(Linea, 16) = LoadResString(C_PUBLIC_FUNCTION) Then 'PUBLIC FUNCTION
                                    EndGeneral = True
                                    nLinea = 1
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaPublicFunction(oTrv, xproyecto)
                                ElseIf InStr(Linea, LoadResString(C_FUNCTION)) And Left$(Linea, 9) = LoadResString(C_FUNCTION) Then        'FUNCTION
                                    EndGeneral = True
                                    nLinea = 1
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaFunction(oTrv, xproyecto)
                                ElseIf InStr(Linea, LoadResString(C_FRIEND_FUNCTION)) And Left$(Linea, 16) = LoadResString(C_FRIEND_FUNCTION) Then   'FRIEND FUNCTION
                                    EndGeneral = True
                                    nLinea = 1
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaFunction(oTrv, xproyecto)
                                ElseIf InStr(Linea, LoadResString(C_END_FUNCTION)) Then          'END FUNCTION
                                    If Left$(Linea, 12) = LoadResString(C_END_FUNCTION) Then
                                        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                        Call FinalizarSub(xproyecto)
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_CONST)) Or InStr(Linea, LoadResString(C_CONST)) Then     'CONSTANTES
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                        nLinea = 1
                                    End If
                                    Call AnalizaPrivateConst(oTrv, xproyecto)
                                ElseIf InStr(Linea, LoadResString(C_PUBLIC_CONST)) Or InStr(Linea, LoadResString(C_GLOBAL_CONST)) Then  'CONSTANTES
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                        nLinea = 1
                                    End If
                                    Call AnalizaPublicConst(oTrv, xproyecto)
                                ElseIf Left$(Linea, 4) = "Type" Or Left$(Linea, 12) = "Private Type" Or Left$(Linea, 11) = "Public Type" Then 'TIPOS
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                        nLinea = 1
                                    End If
                                    Call AnalizaType(oTrv, xproyecto)
                                ElseIf InStr(Linea, LoadResString(C_END_TYPE)) Then   'FIN TIPOS
                                    StartTypes = False
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_ENUM)) Or InStr(Linea, LoadResString(C_PUBLIC_ENUM)) Or InStr(Linea, LoadResString(C_ENUM)) Then  'ENUMERACIONES
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                        nLinea = 1
                                    End If
                                    Call AnalizaEnumeracion(oTrv, xproyecto)
                                ElseIf InStr(Linea, LoadResString(C_END_ENUM)) Then   'FIN ENUM
                                    StartEnum = False
                                ElseIf InStr(Linea, LoadResString(C_PROP_PRIVATE_GET)) Or InStr(Linea, LoadResString(C_PROP_PRIVATE_LET)) Or InStr(Linea, LoadResString(C_PROP_PRIVATE_SET)) Then  'PROPIEDADES
                                    EndGeneral = True
                                    nLinea = 1
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaPropiedad(oTrv, xproyecto, xTotalesProyecto)
                                ElseIf InStr(Linea, LoadResString(C_PROP_PUBLIC_GET)) Or InStr(Linea, LoadResString(C_PROP_PUBLIC_LET)) Or InStr(Linea, LoadResString(C_PROP_PUBLIC_SET)) Then  'PROPIEDADES
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    EndGeneral = True
                                    nLinea = 1
                                    Call AnalizaPropiedad(oTrv, xproyecto, xTotalesProyecto)
                                ElseIf Left$(Linea, 15) = "Friend Property" Then
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    EndGeneral = True
                                    nLinea = 1
                                    Call AnalizaPropiedad(oTrv, xproyecto, xTotalesProyecto)
                                ElseIf Left$(Linea, 12) = "End Property" Then
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call FinalizarSub(xproyecto)
                                ElseIf InStr(Linea, LoadResString(C_EVENTO)) Or InStr(Linea, LoadResString(C_PUBLIC_EVENT)) Then  'EVENTO
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                        nLinea = 1
                                    End If
                                    Call AnalizaEvento(oTrv, xproyecto)
                                ElseIf InStr(Linea, LoadResString(C_DIM)) Or InStr(Linea, LoadResString(C_PRIVATE)) Or InStr(Linea, LoadResString(C_PUBLIC)) Or InStr(Linea, LoadResString(C_GLOBAL)) Or InStr(Linea, LoadResString(C_STATIC)) Then 'VARIABLES
                                    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                        nLinea = 1
                                    End If
                                    Call AnalizaDim(oTrv, xproyecto) 'VARIABLES
                                ElseIf Left$(Linea, 20) = "Attribute " & LoadResString(C_VBNAME) Then 'NOMBRE DEL OBJETO
                                    For j = Len(Linea) To 1 Step -1
                                        If Mid$(Linea, j, 1) = "=" Then
                                            sNombre = Mid$(Linea, j + 2)
                                            sNombre = Left$(sNombre, Len(sNombre) - 1)
                                            sNombre = Mid$(sNombre, 2)
                                            Exit For
                                        End If
                                    Next j
                                    xproyecto.aArchivos(k).ObjectName = sNombre
                                ElseIf InStr(Linea, Trim$(LoadResString(C_BEGIN))) Then
                                    If MyInstr(Linea, Trim$(LoadResString(C_BEGIN))) Then       'NOMBRE DEL CONTROL
                                        Call AnalizaNombreControl(oTrv, xproyecto, xTotalesProyecto)
                                    End If
                                End If
                                                                                                
                                'guardar los elementos del tipo
                                If StartTypes Then
                                    Call DeterminaElementosTipos(oTrv, xproyecto)
                                End If
                                
                                'guardar los elementos de la enumeracion
                                If StartEnum Then
                                    Call DeterminaElementosEnumeracion(oTrv, xproyecto)
                                End If
                                                                
                                'chequear espacios en blancos/comentarios/acumular
                                Call ChequeaLineaDeRutina(xproyecto, FlagLinea)
                            Else
                                'chequear espacios en blancos/comentarios/acumular
                                Call ChequeaLineaDeRutina(xproyecto, FlagLinea)
                            End If
                        Else
                            Call ChequeaLineaDeRutina(xproyecto, FlagLinea)
                        End If
                        
                        If (i Mod 100) = 0 Then InvalidateRect oTrv.hWnd, 0&, 0&
                        i = i + 1
                    Loop
                Close #nFreeFile
            Else
                MsgBox "Error al abrir el archivo : " & xproyecto.aArchivos(k).PathFisico, vbCritical
            End If
            
            Call AcumuladoresParciales(oTrv, xproyecto)
            
            Call AcumularTotalesParciales(xTotalesProyecto, xproyecto, k, apri, apub, cpri, cpub, epri, epub, fpri, fpub, spri, spub, tpri, tpub, vpri, vpub)
        End If
    Next k
    
    InvalidateRect oTrv.hWnd, 0&, 0&
    
End Sub

'almacena los elementos de la enumeracion
Private Sub DeterminaElementosEnumeracion(oTrv As TreeView, xproyecto As eProyecto)

    Dim Enumeracion As String
    Dim total As Integer
    Dim KeyNode As String
    Dim Elemento As String
    
    Enumeracion = Trim$(LineaOrigen)
    
    If InStr(Linea, LoadResString(C_PRIVATE_ENUM)) Then
        Exit Sub
    ElseIf InStr(Linea, LoadResString(C_PUBLIC_ENUM)) Then
        Exit Sub
    ElseIf InStr(Linea, LoadResString(C_ENUM)) Then
        Exit Sub
    ElseIf Left$(Enumeracion, 1) = "'" Then Exit Sub
        Exit Sub
    End If
    
    Enumeracion = NombreX(LineaOrigen)
    
    total = UBound(xproyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos) + 1
    
    ReDim Preserve xproyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total)
    
    If InStr(Enumeracion, "=") Then
        xproyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Nombre = Trim$(Left$(Enumeracion, InStr(1, Enumeracion, "=") - 1))
        xproyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Valor = Trim$(Mid$(Enumeracion, InStr(Enumeracion, "=") + 1))
    Else
        xproyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Nombre = Enumeracion
        xproyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Valor = ""
    End If
    
    Elemento = Enumeracion
    xproyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Estado = NOCHEQUEADO
    
    xproyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Linea = nLinea
    xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
    
    ENUMECH = ENUMECH + 1
    nLinea = nLinea + 1
    
    '*******
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        KeyNode = "FENUMCH" & ENUMECH
        Call oTrv.Nodes.Add("FENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        KeyNode = "BENUMCH" & ENUMECH
        Call oTrv.Nodes.Add("BENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        KeyNode = "CENUMCH" & ENUMECH
        Call oTrv.Nodes.Add("CENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        KeyNode = "KENUMCH" & ENUMECH
        Call oTrv.Nodes.Add("KENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        KeyNode = "PENUMCH" & ENUME
        Call oTrv.Nodes.Add("PENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        KeyNode = "DENUMCH" & ENUME
        Call oTrv.Nodes.Add("DENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    End If
    
    xproyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).KeyNode = KeyNode
            
End Sub


'almacena los elementos del tipo y los guarda en el arreglo
Private Sub DeterminaElementosTipos(oTrv As TreeView, xproyecto As eProyecto)

    Dim Elemento As String
    Dim total As Integer
    Dim TipoVb As String
    Dim KeyNode As String
    
    Elemento = Linea
    
    If Left$(Tipo, 4) = LoadResString(C_TYPE) Then Exit Sub
    If Left$(Tipo, 11) = LoadResString(C_PUBLIC_TYPE) Then Exit Sub
    If Left$(Tipo, 12) = LoadResString(C_PRIVATE_TYPE) Then Exit Sub
    If Left$(Elemento, 1) = "'" Then
        Exit Sub
    End If
    
    Elemento = NombreX(Elemento)
    
    If InStr(Elemento, LoadResString(C_AS)) Then
        total = UBound(xproyecto.aArchivos(k).aTipos(t - 1).aElementos()) + 1
        
        ReDim Preserve xproyecto.aArchivos(k).aTipos(t - 1).aElementos(total)
        
        xproyecto.aArchivos(k).aTipos(t - 1).aElementos(total).Nombre = Left$(Elemento, InStr(Elemento, LoadResString(C_AS)) - 1)
        xproyecto.aArchivos(k).aTipos(t - 1).aElementos(total).Tipo = DeterminaTipoVariable(Elemento, False, TipoVb)
        xproyecto.aArchivos(k).aTipos(t - 1).aElementos(total).Estado = NOCHEQUEADO
        xproyecto.aArchivos(k).aTipos(t - 1).aElementos(total).Linea = nLinea
        xproyecto.aArchivos(k).NumeroDeLineas = xproyecto.aArchivos(k).NumeroDeLineas + 1
        
        TYPOCH = TYPOCH + 1
        nLinea = nLinea + 1
        
        'agregar los elementos del tipo al arbol
        '*****
        If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            KeyNode = "FTIPCH" & TYPOCH
            Call oTrv.Nodes.Add("FTIP" & TYPO - 1, tvwChild, KeyNode, Elemento, C_ICONO_TIPOS, C_ICONO_TIPOS)
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            KeyNode = "BTIPCH" & TYPOCH
            Call oTrv.Nodes.Add("BTIP" & TYPO - 1, tvwChild, KeyNode, Elemento, C_ICONO_TIPOS, C_ICONO_TIPOS)
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            KeyNode = "CTIPCH" & TYPOCH
            Call oTrv.Nodes.Add("CTIP" & TYPO - 1, tvwChild, KeyNode, Elemento, C_ICONO_TIPOS, C_ICONO_TIPOS)
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            KeyNode = "KTIPCH" & TYPOCH
            Call oTrv.Nodes.Add("KTIP" & TYPO - 1, tvwChild, KeyNode, Elemento, C_ICONO_TIPOS, C_ICONO_TIPOS)
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            KeyNode = "PTIPCH" & TYPOCH
            Call oTrv.Nodes.Add("PTIP" & TYPO - 1, tvwChild, KeyNode, Elemento, C_ICONO_TIPOS, C_ICONO_TIPOS)
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            KeyNode = "DTIPCH" & TYPOCH
            Call oTrv.Nodes.Add("DTIP" & TYPO - 1, tvwChild, KeyNode, Elemento, C_ICONO_TIPOS, C_ICONO_TIPOS)
        End If
        
        xproyecto.aArchivos(k).aTipos(t - 1).aElementos(total).KeyNode = KeyNode
        
        '*****
    End If
    
End Sub


'determina el tipo de variable
Private Function DeterminaTipoVariable(ByVal Variable As String, Predefinido As Boolean, _
                                       ByVal TipoVb As String)

    Dim TipoDefinido As String
    
    Predefinido = False
    
    If InStr(Variable, LoadResString(C_AS)) Then
        TipoDefinido = Mid$(Variable, InStr(Variable, LoadResString(C_AS)) + 1)
        
        If InStr(TipoDefinido, "*") > 0 Then
            TipoDefinido = Left$(TipoDefinido, InStr(1, TipoDefinido, "*") - 1)
            TipoDefinido = Trim$(Mid$(TipoDefinido, 4))
        ElseIf InStr(TipoDefinido, "'") > 0 Then
            TipoDefinido = Left$(TipoDefinido, InStr(1, TipoDefinido, "'") - 1)
            TipoDefinido = Trim$(Mid$(TipoDefinido, 4))
        Else
            TipoDefinido = Trim$(Mid$(TipoDefinido, 4))
        End If
        
        If InStr(1, TipoDefinido, " ") <> 0 Then    'POR SI VIENE NEW
            TipoDefinido = Mid$(TipoDefinido, InStr(1, TipoDefinido, " ") + 1)
        End If
    Else
        TipoDefinido = "Variant"
        TipoVb = ""
        Predefinido = True
    End If
    
    DeterminaTipoVariable = TipoDefinido
    
End Function


'devuelve el nombre de la variable sin comentarios
Private Function NombreX(ByVal Variable As String) As String

    Dim ret As String
    
    ret = Variable
    
    If InStr(1, Variable, "'") Then
        ret = Trim$(Left$(Variable, InStr(1, Variable, "'") - 1))
    End If
    
    NombreX = ret
    
End Function

'devuelve lo que esta a la derecha de la funcion
Private Function RetornoFuncion(ByVal Funcion As String) As String

    Dim ret As String
    Dim k As Integer
    
    For k = Len(Funcion) To 1 Step -1
        If Mid$(Funcion, k, 1) = " " Then
            ret = Mid$(Funcion, k + 1)
            Exit For
        End If
    Next k
    
    RetornoFuncion = ret
    
End Function

'setear nombre de la rutina con los parametros analizados
Private Sub SetearNombreRutina(oTrv As TreeView, xproyecto As eProyecto)

    Dim j As Integer
    Dim FinRutina As String
            
    If UBound(xproyecto.aArchivos(k).aRutinas(r - 1).Aparams()) > 0 Then
        xproyecto.aArchivos(k).aRutinas(r - 1).Nombre = ""
        'setear nombre de la sub/funcion/propiedad
        For j = 1 To UBound(xproyecto.aArchivos(k).aRutinas(r - 1).Aparams())
            If Right$(Trim$(xproyecto.aArchivos(k).aRutinas(r - 1).Aparams(j).Glosa), 1) <> "(" Then
                xproyecto.aArchivos(k).aRutinas(r - 1).Nombre = xproyecto.aArchivos(k).aRutinas(r - 1).Nombre & xproyecto.aArchivos(k).aRutinas(r - 1).Aparams(j).Glosa & " , "
            Else
                xproyecto.aArchivos(k).aRutinas(r - 1).Nombre = xproyecto.aArchivos(k).aRutinas(r - 1).Nombre & xproyecto.aArchivos(k).aRutinas(r - 1).Aparams(j).Glosa
            End If
        Next j
        
        xproyecto.aArchivos(k).aRutinas(r - 1).Nombre = Left$(xproyecto.aArchivos(k).aRutinas(r - 1).Nombre, Len(xproyecto.aArchivos(k).aRutinas(r - 1).Nombre) - 3)
        
        If InStr(1, Linea, ")") > 0 Then
            FinRutina = Mid$(Linea, InStr(1, Linea, ")"))
        Else
            FinRutina = Linea
        End If
        
        xproyecto.aArchivos(k).aRutinas(r - 1).Nombre = xproyecto.aArchivos(k).aRutinas(r - 1).Nombre & FinRutina
        
        If Right$(xproyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
            xproyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
        Else
            xproyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
            xproyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Funcion)
        End If
    End If
                                
End Sub

'determina si la sub es un evento de un control
Private Sub DeterminaEventosControles(oTrv As TreeView, xproyecto As eProyecto)
    
    Dim j As Integer
    Dim i As Integer
    Dim Evento As String
    Dim sControl As String
    Dim sEventos As String
    
    For k = 1 To UBound(xproyecto.aArchivos)
        If xproyecto.aArchivos(k).Explorar Then
            If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Or _
               xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Or _
               xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                'MsgBox xproyecto.aArchivos(k).PathFisico
                For i = 1 To UBound(xproyecto.aArchivos(k).aRutinas)
                    If xproyecto.aArchivos(k).aRutinas(i).Tipo = TIPO_SUB Then
                        If Left$(xproyecto.aArchivos(k).aRutinas(i).Nombre, 12) = LoadResString(C_PRIVATE_SUB) Then
                            Evento = Mid$(xproyecto.aArchivos(k).aRutinas(i).Nombre, 13)
                        ElseIf Left$(xproyecto.aArchivos(k).aRutinas(i).Nombre, 11) = LoadResString(C_PUBLIC_SUB) Then
                            Evento = Mid$(xproyecto.aArchivos(k).aRutinas(i).Nombre, 12)
                        ElseIf Left$(xproyecto.aArchivos(k).aRutinas(i).Nombre, 4) = LoadResString(C_SUB) Then
                            Evento = Mid$(xproyecto.aArchivos(k).aRutinas(i).Nombre, 5)
                        End If
                        
                        If InStr(Evento, "_") Then
                            Evento = Left$(Evento, InStr(1, Evento, "(") - 1)
                            For j = Len(Evento) To 1 Step -1
                                If Mid$(Evento, j, 1) = "_" Then
                                    sControl = Left$(Evento, j - 1)
                                    Evento = Mid$(Evento, j + 1)
                                    Exit For
                                End If
                            Next j
                            
                            For j = 1 To UBound(xproyecto.aArchivos(k).aControles)
                                If Trim$(UCase$(xproyecto.aArchivos(k).aControles(j).Nombre)) = Trim$(UCase$(sControl)) Then
                                    xproyecto.aArchivos(k).aControles(j).Eventos = xproyecto.aArchivos(k).aControles(j).Eventos & Evento & " , "
                                    Exit For
                                End If
                            Next j
                        End If
                    End If
                Next i
                
                'limpiar la , final
                For j = 1 To UBound(xproyecto.aArchivos(k).aControles)
                    If Right$(xproyecto.aArchivos(k).aControles(j).Eventos, 2) = ", " Then
                        xproyecto.aArchivos(k).aControles(j).Eventos = Left$(xproyecto.aArchivos(k).aControles(j).Eventos, Len(xproyecto.aArchivos(k).aControles(j).Eventos) - 3)
                    End If
                Next j
            End If
        End If
    Next k
    
End Sub

'determina el tipo de xproyecto
Private Function DeterminaTipoDeproyecto(oTrv As TreeView, xproyecto As eProyecto, ByVal Archivo As String) As Boolean

    On Local Error GoTo ErrorDeterminaTipoDexproyecto
    
    Dim ret As Boolean
    Dim Icono As Integer
    Dim sxproyecto As String
    Dim Linea As String
    Dim sNombreArchivo As String
    Dim nFreeFile As Long
    
    Icono = C_ICONO_PROYECTO
        
    sNombreArchivo = VBArchivoSinPath(Archivo)
    
    nFreeFile = FreeFile
    
    ret = True
    
    xproyecto.TipoProyecto = PRO_TIPO_EXE
    xproyecto.Version = 0
    
    Open Archivo For Input Shared As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, Linea
            If Left$(Linea, 4) = "Type" Then
                If MyInstr(Linea, "Exe") Then
                    Icono = C_ICONO_PROYECTO
                    xproyecto.TipoProyecto = PRO_TIPO_EXE
                    xproyecto.Icono = Icono
                ElseIf MyInstr(Linea, "OleExe") Then
                    Icono = C_ICONO_ACTIVEX_EXE
                    xproyecto.TipoProyecto = PRO_TIPO_EXE_X
                    xproyecto.Icono = Icono
                ElseIf MyInstr(Linea, "Control") Then
                    Icono = C_ICONO_OCX
                    xproyecto.TipoProyecto = PRO_TIPO_OCX
                    xproyecto.Icono = Icono
                ElseIf MyInstr(Linea, "OleDll") Then
                    Icono = C_ICONO_DLL
                    xproyecto.TipoProyecto = PRO_TIPO_DLL
                    xproyecto.Icono = Icono
                End If
            ElseIf Left$(Linea, 4) = "Name" Then
                sxproyecto = Mid$(Linea, 6)
                sxproyecto = Mid$(sxproyecto, 2)
                sxproyecto = Left$(sxproyecto, Len(sxproyecto) - 1)
                xproyecto.Nombre = sxproyecto
                xproyecto.Archivo = sNombreArchivo
            End If
        Loop
    Close #nFreeFile
            
    'para versiones de visual basic que no tienen el name
    If xproyecto.TipoProyecto = PRO_TIPO_NONE Then
        xproyecto.TipoProyecto = PRO_TIPO_EXE
        xproyecto.Icono = C_ICONO_PROYECTO
        xproyecto.Version = 3
    End If
    
    If xproyecto.Nombre = "" Then
        xproyecto.Nombre = Left$(sNombreArchivo, InStr(1, sNombreArchivo, ".") - 1)
        xproyecto.Archivo = sNombreArchivo
        xproyecto.Version = 3
    End If
    
    GoTo SalirDeterminaTipoDexproyecto
    
ErrorDeterminaTipoDexproyecto:
    ret = False
    MsgBox "DeterminaTipoDexproyecto : " & Err & " " & Error$, vbCritical
    Resume SalirDeterminaTipoDexproyecto
    
SalirDeterminaTipoDexproyecto:
    DeterminaTipoDeproyecto = ret
    Err = 0
    
End Function
Public Sub EnabledControls(ByVal frm As Form, ByVal bEnabled As Boolean)

    Dim k As Integer
    Dim c As Integer
    Dim oControl As Control
    
    With frm
        For k = 1 To .Controls.Count - 1
            If TypeOf .Controls(k) Is Menu Then
                Set oControl = .Controls(k)
                
                If oControl.Caption <> "-" Then
                    oControl.Enabled = bEnabled
                End If
            ElseIf TypeOf .Controls(k) Is Toolbar Then
                Set oControl = .Controls(k)
                
                For c = 1 To oControl.Buttons.Count
                    oControl.Buttons(c).Enabled = bEnabled
                Next c
            End If
        Next k
    End With
    
    Set oControl = Nothing
    
End Sub
'finaliza la rutina y determina el numero de lineas
Private Sub FinalizarSub(xproyecto As eProyecto)

    Call GrabaLineaDeRutina(xproyecto)
    bEndSub = True
    'Close #FreeSub
    aru = 1
    
    StartRutinas = False
    
    'numero de lineas de la rutina
    'totalineas - comentarios - blancos
    xproyecto.aArchivos(k).aRutinas(r - 1).NumeroDeLineas = _
    xproyecto.aArchivos(k).aRutinas(r - 1).TotalLineas - _
    xproyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios - _
    xproyecto.aArchivos(k).aRutinas(r - 1).NumeroDeBlancos
            
    Err = 0
    
End Sub

'guardar la rutina en el arreglo de rutinas
Private Sub GrabaLineaDeRutina(xproyecto As eProyecto)
    
    If UBound(Arr_Paso) > 0 Then
        Dim rp As Integer
        For rp = 1 To UBound(Arr_Paso)
            ReDim Preserve xproyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru)
            xproyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).Codigo = Arr_Paso(rp)
            aru = aru + 1
        Next rp
        ReDim Arr_Paso(0)
    Else
        ReDim Preserve xproyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru)
        xproyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).Codigo = LineaOrigen
        xproyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).Linea = nLinea
        aru = aru + 1
    End If
        
End Sub

Public Sub HelpCarga(ByVal Ayuda As String)
    frmMain.stbMain.Panels(1).Text = Ayuda
End Sub

Private Sub InicializarVariables(xproyecto As eProyecto)

    nFreeFile = FreeFile
    
    PROC = 1
    Func = 1
    api = 1
    Cons = 1
    TYPO = 1
    TYPOCH = 0
    VARY = 1
    ENUME = 1
    ENUMECH = 0
    ARRAYY = 1
    VARYPROC = 1
    VARYPROP = 1
    NPROP = 1
    NEVENTO = 1
    PRIFUN = 1
    PUBFUN = 1
    PRISUB = 1
    PUBSUB = 1
    
    frmMain.pgbStatus.Max = UBound(xproyecto.aArchivos) + 1
    frmMain.pgbStatus.Value = 1
    
    xproyecto.FileSize = VBGetFileSize(xproyecto.PathFisico)
    
    'gsTempPath = VBGetTempPath()
    
End Sub

Private Sub InicializarVariablesArchivos(xproyecto As eProyecto)

    r = 1
    i = 1
    c = 1
    t = 1
    vr = 1
    vpro = 1
    v = 1
    e = 1
    ap = 1
    a = 1
    f = 1
    s = 1
    ge = 1
    
    spri = 1
    spub = 1
    fpri = 1
    fpub = 1
    cpri = 1
    cpub = 1
    epri = 1
    epub = 1
    tpri = 1
    tpub = 1
    vpri = 1
    vpub = 1
    apri = 1
    apub = 1
    aru = 1
    apro = 1
    ca = 1
    prop = 1
    even = 1
    
    NumeroDeLineas = 1
    
    StartRutinas = False
    StartHeader = False
    StartGeneral = False
    StartTypes = False
    StartEnum = False
        
    EndHeader = True
    EndGeneral = False
    
    bSub = False
    bSubPub = False
    bSubPri = False
    
    bFun = False
    bFunPub = False
    bFunPri = False
    
    bApi = False
    bCon = False
    bTipo = False
    bVariables = False
    bEnumeracion = False
    bArray = False
    bEndSub = True
    bPropiedades = False
    bEventos = False
    bEndProp = True
    
    frmMain.pgbStatus.Value = k
    frmMain.stbMain.Panels(2).Text = k & " de " & UBound(xproyecto.aArchivos)
    frmMain.stbMain.Panels(4).Text = Round(k * 100 / UBound(xproyecto.aArchivos), 0) & " %"
    Call HelpCarga("Leyendo : " & xproyecto.aArchivos(k).Nombre)
        
    xproyecto.aArchivos(k).OptionExplicit = False
    
    xproyecto.aArchivos(k).nArray = 0
    xproyecto.aArchivos(k).nConstantes = 0
    xproyecto.aArchivos(k).nEnumeraciones = 0
    xproyecto.aArchivos(k).nTipos = 0
    xproyecto.aArchivos(k).nVariables = 0
    xproyecto.aArchivos(k).nTipoApi = 0
    xproyecto.aArchivos(k).nTipoFun = 0
    xproyecto.aArchivos(k).nTipoSub = 0
    xproyecto.aArchivos(k).nTipoApi = 0
    xproyecto.aArchivos(k).NumeroDeLineas = 0
    xproyecto.aArchivos(k).nControles = 0
    xproyecto.aArchivos(k).nEventos = 0
    xproyecto.aArchivos(k).nPropiedades = 0
    xproyecto.aArchivos(k).nPropertyGet = 0
    xproyecto.aArchivos(k).nPropertyLet = 0
    xproyecto.aArchivos(k).nPropertySet = 0
    xproyecto.aArchivos(k).NumeroDeLineasComentario = 0
    xproyecto.aArchivos(k).NumeroDeLineasEnBlanco = 0
    xproyecto.aArchivos(k).MiembrosPublicos = 0
    xproyecto.aArchivos(k).MiembrosPrivados = 0
    
    ReDim xproyecto.aArchivos(k).aGeneral(0)
    ReDim xproyecto.aArchivos(k).aTipoVariable(0)
    ReDim xproyecto.aArchivos(k).aArray(0)
    ReDim xproyecto.aArchivos(k).aConstantes(0)
    ReDim xproyecto.aArchivos(k).aEnumeraciones(0)
    ReDim xproyecto.aArchivos(k).aRutinas(0)
    ReDim xproyecto.aArchivos(k).aRutinas(0).aVariables(0)
    ReDim xproyecto.aArchivos(k).aRutinas(0).aRVariables(0)
    ReDim xproyecto.aArchivos(k).aTipos(0)
    ReDim xproyecto.aArchivos(k).aVariables(0)
    ReDim xproyecto.aArchivos(k).aTipoVariable(0)
    ReDim xproyecto.aArchivos(k).aControles(0)
    ReDim xproyecto.aArchivos(k).aApis(0)
    ReDim xproyecto.aArchivos(k).aEventos(0)
    
    xproyecto.aArchivos(k).FileSize = VBGetFileSize(xproyecto.aArchivos(k).PathFisico)
        
    LineaPaso = ""
    
End Sub
'LIMPIAR TOTALES GENERALES
Private Sub LimpiarTotales(xTotalesProyecto As eTotalesProyecto)

    xTotalesProyecto.TotalVariables = 0
    xTotalesProyecto.TotalConstantes = 0
    xTotalesProyecto.TotalEnumeraciones = 0
    xTotalesProyecto.TotalArray = 0
    xTotalesProyecto.TotalTipos = 0
    xTotalesProyecto.TotalSubs = 0
    xTotalesProyecto.TotalFunciones = 0
    xTotalesProyecto.TotalApi = 0
    xTotalesProyecto.TotalEventos = 0
    xTotalesProyecto.TotalPropiedades = 0
    xTotalesProyecto.TotalArrayPrivadas = 0
    xTotalesProyecto.TotalArrayPublicas = 0
    xTotalesProyecto.TotalConstantesPrivadas = 0
    xTotalesProyecto.TotalConstantesPublicas = 0
    xTotalesProyecto.TotalEnumeracionesPrivadas = 0
    xTotalesProyecto.TotalEnumeracionesPublicas = 0
    xTotalesProyecto.TotalFuncionesPrivadas = 0
    xTotalesProyecto.TotalFuncionesPublicas = 0
    xTotalesProyecto.TotalSubsPrivadas = 0
    xTotalesProyecto.TotalSubsPublicas = 0
    xTotalesProyecto.TotalTiposPrivadas = 0
    xTotalesProyecto.TotalTiposPublicas = 0
    xTotalesProyecto.TotalVariablesPrivadas = 0
    xTotalesProyecto.TotalVariablesPublicas = 0
    xTotalesProyecto.TotalLineasDeCodigo = 0
    xTotalesProyecto.TotalLineasDeComentarios = 0
    xTotalesProyecto.TotalLineasEnBlancos = 0
    xTotalesProyecto.TotalPropertyGets = 0
    xTotalesProyecto.TotalPropertyLets = 0
    xTotalesProyecto.TotalPropertySets = 0
    xTotalesProyecto.TotalControles = 0
    
End Sub



Private Function NombreArchivo(ByVal sLinea As String, ByVal Leer As Integer) As String

    Dim k As Integer
    Dim ret As String
    Dim Inicio As Integer
    
    Inicio = 0
    
    If Leer = 1 Then        'REFERENCIAS
        For k = Len(sLinea) To 1 Step -1
            If Mid$(sLinea, k, 1) = "#" Then
                If Inicio = 0 Then
                    Inicio = k
                Else
                    ret = Mid$(sLinea, k + 1, Inicio - (k + 1))
                    Exit For
                End If
            End If
        Next k
    ElseIf Leer = 2 Then    'CONTROLES
        For k = Len(sLinea) To 1 Step -1
            If Mid$(sLinea, k, 1) = ";" Then
                Inicio = k
                ret = Trim$(Mid$(sLinea, Inicio + 1))
                Exit For
            End If
        Next k
    End If
    
    NombreArchivo = ret
    
End Function

Public Function PathArchivo(ByVal Archivo As String) As String

    Dim k As Integer
    
    Dim ret As String
    
    ret = ""
    
    For k = Len(Archivo) To 1 Step -1
        If Mid$(Archivo, k, 1) = "\" Then
            ret = Mid$(Archivo, 1, k)
            Exit For
        End If
    Next k
    
    PathArchivo = ret
    
End Function

'procesar los parametros que vienen
Private Sub ProcesarParametros(oTrv As TreeView, xproyecto As eProyecto)

    Dim params As Integer
    Dim StartParam As Integer
    Dim sParam As String
    Dim Parametro As String
    Dim TipoParametro As String
    Dim Inicio As Integer
    Dim Fin As Boolean
    Dim Fin2 As Boolean
    Dim Nombre As String
    Dim j As Integer
    Dim Glosa As String
    Dim PorValor As Boolean
    Dim ArrayParam As Boolean
    
    StartParam = 0
    sParam = Linea
            
    Do  'ciclar por los parametros
        ArrayParam = False
        If InStr(1, sParam, ",") <> 0 Then
            Parametro = Trim$(Left$(sParam, InStr(1, sParam, ",") - 1))
            
            Inicio = InStr(1, sParam, ",") + 1
            sParam = Trim$(Mid$(sParam, Inicio))
            Fin = False
        ElseIf InStr(sParam, ")") > 0 Then
            Parametro = Trim$(Left$(sParam, InStr(1, sParam, ")") - 1))
            
            If Right$(Parametro, 1) = "(" Then
                sParam = Mid$(sParam, InStr(1, sParam, ")"))
                
                If InStr(2, sParam, ")") > 0 Then
                    Parametro = Parametro & Left$(sParam, InStr(2, sParam, ")") - 1)
                    sParam = Mid$(sParam, InStr(2, sParam, ")") + 1)
                ElseIf Right$(sParam, 1) = "_" Then
                    Parametro = Trim$(Parametro & Left$(sParam, Len(sParam) - 1))
                    sParam = ""
                End If
                ArrayParam = True
            End If
            
            Inicio = InStr(1, Parametro, ",") + 1
            Fin = False
            Fin2 = True
        Else
            Fin = True
        End If
        
        If Parametro <> "" Then
            If Not Fin Or UBound(xproyecto.aArchivos(k).aRutinas(r - 1).Aparams()) = 0 Then
                params = UBound(xproyecto.aArchivos(k).aRutinas(r - 1).Aparams()) + 1
                ReDim Preserve xproyecto.aArchivos(k).aRutinas(r - 1).Aparams(params)
                
                PorValor = False
                'primer parametro puede venir la glosa del sub/fun/propiedad
                If Parametro <> "" Then
                    Glosa = Parametro
                    If params = 1 Then
                        If Not ArrayParam Then
                            If InStr(Parametro, "(") = 0 Then
                                Parametro = Mid$(Parametro, InStr(Parametro, "(") + 1)
                            End If
                        Else
                            If InStr(Parametro, "(") = 0 Then
                                Parametro = Trim$(Mid$(Parametro, InStr(Parametro, ")") + 1))
                            End If
                        End If
                    End If
                Else
                    Parametro = Left$(sParam, Len(sParam) - 1)
                    Glosa = Trim$(Parametro)
                End If
                
                'desglosar el parametro
                If InStr(Parametro, "ByVal") <> 0 Or InStr(Parametro, "ByRef") <> 0 Then
                    
                    'determinar si viene por valor o x referencia
                    If InStr(Parametro, "ByVal") <> 0 Then
                        PorValor = True
                    End If
                    
                    If Left$(Parametro, 14) = "Optional ByVal" Then
                        Parametro = Mid$(Parametro, 16)
                    ElseIf Left$(Parametro, 14) = "Optional ByRef" Then
                        Parametro = Mid$(Parametro, 16)
                    ElseIf Left$(Parametro, 5) = "ByVal" Then
                        Parametro = Mid$(Parametro, 7)
                    ElseIf Left$(Parametro, 5) = "ByRef" Then
                        Parametro = Mid$(Parametro, 7)
                    End If
                    
                    If InStr(Parametro, LoadResString(C_AS)) <> 0 Then
                        Nombre = Left$(Parametro, InStr(Parametro, LoadResString(C_AS)) - 1)
                        TipoParametro = Mid$(Parametro, InStr(Parametro, LoadResString(C_AS)) + 4)
                    Else
                        Parametro = Trim$(Parametro)
                        If InStr(1, Parametro, "=") Then
                            Parametro = Trim$(Left$(Parametro, InStr(1, Parametro, "=") - 1))
                        End If
                        
                        If Not BasicOldStyle(Parametro) Then
                            Nombre = Parametro
                            TipoParametro = "Variant"
                        Else
                            Nombre = Left$(Parametro, Len(Parametro) - 1)
                            TipoParametro = Right$(Parametro, 1)
                        End If
                    End If
                Else
                    If InStr(Parametro, LoadResString(C_AS)) <> 0 Then
                        If InStr(Parametro, "Optional") = 0 Then
                            Nombre = Left$(Parametro, InStr(Parametro, LoadResString(C_AS)) - 1)
                            TipoParametro = Mid$(Parametro, InStr(Parametro, LoadResString(C_AS)) + 4)
                        Else
                            Parametro = Mid$(Parametro, 10)
                            Nombre = Left$(Parametro, InStr(Parametro, LoadResString(C_AS)) - 1)
                            
                            If InStr(Parametro, "=") = 0 Then
                                TipoParametro = Mid$(Parametro, InStr(Parametro, LoadResString(C_AS)) + 4)
                            Else
                                Parametro = Mid$(Parametro, InStr(Parametro, LoadResString(C_AS)) + 4)
                                TipoParametro = Trim$(Left$(Parametro, InStr(1, Parametro, "=") - 1))
                            End If
                        End If
                    Else
                        If Left$(Parametro, 14) = "Optional ByVal" Then
                            Parametro = Mid$(Parametro, 16)
                        ElseIf Left$(Parametro, 14) = "Optional ByRef" Then
                            Parametro = Mid$(Parametro, 16)
                        ElseIf Left$(Parametro, 5) = "ByVal" Then
                            Parametro = Mid$(Parametro, 7)
                        ElseIf Left$(Parametro, 5) = "ByRef" Then
                            Parametro = Mid$(Parametro, 7)
                        ElseIf Left$(Parametro, 8) = "Optional" Then
                            Parametro = Mid$(Parametro, 9)
                        End If
                        
                        If InStr(Parametro, LoadResString(C_AS)) <> 0 Then
                            Nombre = Left$(Parametro, InStr(Parametro, LoadResString(C_AS)) - 1)
                            TipoParametro = Mid$(Parametro, InStr(Parametro, LoadResString(C_AS)) + 4)
                        Else
                            Parametro = Trim$(Parametro)
                            
                            If InStr(1, Parametro, "=") Then
                                Parametro = Trim$(Left$(Parametro, InStr(1, Parametro, "=") - 1))
                            End If
                        
                            If Not BasicOldStyle(Parametro) Then
                                Nombre = Parametro
                                TipoParametro = "Variant"
                            Else
                                Nombre = Left$(Parametro, Len(Parametro) - 1)
                                TipoParametro = Right$(Parametro, 1)
                            End If
                        End If
                    End If
                End If
                
                If InStr(1, TipoParametro, "=") Then
                    TipoParametro = Left$(TipoParametro, InStr(1, TipoParametro, "=") - 2)
                End If
                
                If Left$(Nombre, 10) = "ParamArray" Then
                    Nombre = Mid$(Nombre, 12)
                End If
                
                If InStr(1, Nombre, "(") Then
                    Nombre = Left$(Nombre, Len(Nombre) - 2)
                End If
                
                xproyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).Nombre = Nombre
                xproyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).Glosa = Glosa
                xproyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).TipoParametro = TipoParametro
                xproyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).PorValor = PorValor
                xproyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).BasicStyle = BasicOldStyle(Glosa)
                If Fin2 Then
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
End Sub

'acumular los tipos de variables tanto a nivel global
'como a nivel de rutinas
Private Sub ProcesarTipoDeVariable(oTrv As TreeView, xproyecto As eProyecto, ByVal Variable As String)

    Dim j As Integer
    Dim nTipoVar As Integer
    Dim Found As Boolean
    Dim TipoDefinido As String
    Dim Predefinido As Boolean
    
    TipoDefinido = DeterminaTipoVariable(Variable, Predefinido, "")
                                
    nTipoVar = UBound(xproyecto.aArchivos(k).aTipoVariable())
                            
    Found = False
    For j = 1 To nTipoVar
        If xproyecto.aArchivos(k).aTipoVariable(j).TipoDefinido = TipoDefinido Then
            Found = True
            Exit For
        End If
    Next j
    
    If Not Found Then
        nTipoVar = UBound(xproyecto.aArchivos(k).aTipoVariable()) + 1
        ReDim Preserve xproyecto.aArchivos(k).aTipoVariable(nTipoVar)
        xproyecto.aArchivos(k).aTipoVariable(nTipoVar).Cantidad = 1
        xproyecto.aArchivos(k).aTipoVariable(nTipoVar).TipoDefinido = TipoDefinido
    Else
        xproyecto.aArchivos(k).aTipoVariable(j).Cantidad = _
        xproyecto.aArchivos(k).aTipoVariable(j).Cantidad + 1
    End If
        
    'procesar variables de rutinas
    If StartRutinas Then
        nTipoVar = UBound(xproyecto.aArchivos(k).aRutinas(r - 1).aRVariables)
                            
        Found = False
        For j = 1 To nTipoVar
            If xproyecto.aArchivos(k).aRutinas(r - 1).aRVariables(j).TipoDefinido = TipoDefinido Then
                Found = True
                Exit For
            End If
        Next j
        
        If Not Found Then
            nTipoVar = UBound(xproyecto.aArchivos(k).aRutinas(r - 1).aRVariables()) + 1
            ReDim Preserve xproyecto.aArchivos(k).aRutinas(r - 1).aRVariables(nTipoVar)
            xproyecto.aArchivos(k).aRutinas(r - 1).aRVariables(nTipoVar).Cantidad = 1
            xproyecto.aArchivos(k).aRutinas(r - 1).aRVariables(nTipoVar).TipoDefinido = TipoDefinido
        Else
            xproyecto.aArchivos(k).aRutinas(r - 1).aRVariables(j).Cantidad = _
            xproyecto.aArchivos(k).aRutinas(r - 1).aRVariables(j).Cantidad + 1
        End If
    End If
    
End Sub
'setea los contadores por archivo en el arbol de xproyecto
Private Sub SeteaContadoresAnalisis(oTrv As TreeView, xproyecto As eProyecto)

    Dim k As Long
    Dim j As Long
    
    Dim texto As String
    
    For k = 1 To UBound(xproyecto.aArchivos)
        If xproyecto.aArchivos(k).Explorar Then
            texto = ""
            
            If xproyecto.aArchivos(k).nTipoSub > 0 Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeSub).Text
                texto = texto & "-(" & xproyecto.aArchivos(k).nTipoSub & ")"
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeSub).Text = texto
            End If
            '*
            If xproyecto.aArchivos(k).nTipoFun > 0 Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeFun).Text
                texto = texto & "-(" & xproyecto.aArchivos(k).nTipoFun & ")"
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeFun).Text = texto
            End If
            '*
            If xproyecto.aArchivos(k).nPropiedades > 0 Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeProp).Text
                texto = texto & "-(" & xproyecto.aArchivos(k).nPropiedades & ")"
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeProp).Text = texto
            End If
            '*
            If xproyecto.aArchivos(k).nVariables > 0 Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeVar).Text
                texto = texto & "-(" & xproyecto.aArchivos(k).nVariables & ")"
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeVar).Text = texto
            End If
            '*
            If xproyecto.aArchivos(k).nConstantes > 0 Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeCte).Text
                texto = texto & "-(" & xproyecto.aArchivos(k).nConstantes & ")"
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeCte).Text = texto
            End If
            '*
            If xproyecto.aArchivos(k).nTipos > 0 Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeTipo).Text
                texto = texto & "-(" & xproyecto.aArchivos(k).nTipos & ")"
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeTipo).Text = texto
            End If
            '*
            If xproyecto.aArchivos(k).nEnumeraciones > 0 Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeEnum).Text
                texto = texto & "-(" & xproyecto.aArchivos(k).nEnumeraciones & ")"
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeEnum).Text = texto
            End If
            '*
            If xproyecto.aArchivos(k).nTipoApi > 0 Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeApi).Text
                texto = texto & "-(" & xproyecto.aArchivos(k).nTipoApi & ")"
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeApi).Text = texto
            End If
            '*
            If xproyecto.aArchivos(k).nArray > 0 Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeArr).Text
                texto = texto & "-(" & xproyecto.aArchivos(k).nArray & ")"
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeArr).Text = texto
            End If
            '*
            'MsgBox xproyecto.aArchivos(k).PathFisico
            If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeFrm).Text
                texto = xproyecto.aArchivos(k).ObjectName & " (" & texto & ")"
                
                xproyecto.aArchivos(k).Descripcion = texto
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeFrm).Text = texto
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeBas).Text
                texto = xproyecto.aArchivos(k).ObjectName & " (" & texto & ")"
                
                xproyecto.aArchivos(k).Descripcion = texto
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeBas).Text = texto
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeCls).Text
                texto = xproyecto.aArchivos(k).ObjectName & " (" & texto & ")"
                
                xproyecto.aArchivos(k).Descripcion = texto
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeCls).Text = texto
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeKtl).Text
                texto = xproyecto.aArchivos(k).ObjectName & " (" & texto & ")"
                
                xproyecto.aArchivos(k).Descripcion = texto
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeKtl).Text = texto
            
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodePag).Text
                texto = xproyecto.aArchivos(k).ObjectName & " (" & texto & ")"
                
                xproyecto.aArchivos(k).Descripcion = texto
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodePag).Text = texto
            
            ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeDsr).Text
                texto = xproyecto.aArchivos(k).ObjectName & " (" & texto & ")"
                
                xproyecto.aArchivos(k).Descripcion = texto
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeDsr).Text = texto
            End If
            '*
            'If xproyecto.aArchivos(k).nControles > 0 Then
            '    Texto = otrv.Nodes(xproyecto.aArchivos(k).KeyNodeKtl).Text
            '    Texto = Texto & "-(" & xproyecto.aArchivos(k).nControles & ")"
            '    otrv.Nodes(xproyecto.aArchivos(k).KeyNodeKtl).Text = Texto
            'End If
            
            '*
            If xproyecto.aArchivos(k).nEventos > 0 Then
                texto = oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeEvento).Text
                texto = texto & "-(" & xproyecto.aArchivos(k).nEventos & ")"
                oTrv.Nodes(xproyecto.aArchivos(k).KeyNodeEvento).Text = texto
            End If
        End If
    Next k
    
End Sub

'setear el hijo de la constante en el arbol
Private Sub StartChildConstante(oTrv As TreeView, xproyecto As eProyecto, ByVal NombreConstante As String)

    Dim KeyNode As String
    
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_CONS_FRM & k, tvwChild, "FCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "FCON" & Cons
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_CONS_BAS & k, tvwChild, "BCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "BCON" & Cons
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_CONS_CLS & k, tvwChild, "CCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "CCON" & Cons
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_CONS_CTL & k, tvwChild, "KCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "KCON" & Cons
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_CONS_PAG & k, tvwChild, "PCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "PCON" & Cons
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_CONS_DSR & k, tvwChild, "DCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "DCON" & Cons
    End If
    
    xproyecto.aArchivos(k).aConstantes(c).KeyNode = KeyNode
    
End Sub

'agregar hijo de la variable
Private Sub StartChildDim(oTrv As TreeView, xproyecto As eProyecto)

    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        If StartRutinas Then
            Call oTrv.Nodes.Add(xproyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "FVARPROC" & VARYPROC, xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "FVARPROC" & VARYPROC
        ElseIf Not StartRutinas Then
            Call oTrv.Nodes.Add(C_VAR_FRM & k, tvwChild, "FVAR" & VARY, xproyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aVariables(v).KeyNode = "FVAR" & VARY
        End If
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        If StartRutinas Then
            Call oTrv.Nodes.Add(xproyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "BVARPROC" & VARYPROC, xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "BVARPROC" & VARYPROC
        ElseIf Not StartRutinas Then
            Call oTrv.Nodes.Add(C_VAR_BAS & k, tvwChild, "BVAR" & VARY, xproyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aVariables(v).KeyNode = "BVAR" & VARY
        End If
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        If StartRutinas Then
            Call oTrv.Nodes.Add(xproyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "CVARPROC" & VARYPROC, xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "CVARPROC" & VARYPROC
        ElseIf Not StartRutinas Then
            Call oTrv.Nodes.Add(C_VAR_CLS & k, tvwChild, "CVAR" & VARY, xproyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aVariables(v).KeyNode = "CVAR" & VARY
        End If
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        If StartRutinas Then
            Call oTrv.Nodes.Add(xproyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "KVARPROC" & VARYPROC, xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "KVARPROC" & VARYPROC
        ElseIf Not StartRutinas Then
            Call oTrv.Nodes.Add(C_VAR_CTL & k, tvwChild, "KVAR" & VARY, xproyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aVariables(v).KeyNode = "KVAR" & VARY
        End If
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        If StartRutinas Then
            Call oTrv.Nodes.Add(xproyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "PVARPROC" & VARYPROC, xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "PVARPROC" & VARYPROC
        ElseIf Not StartRutinas Then
            Call oTrv.Nodes.Add(C_VAR_PAG & k, tvwChild, "PVAR" & VARY, xproyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aVariables(v).KeyNode = "PVAR" & VARY
        End If
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        If StartRutinas Then
            Call oTrv.Nodes.Add(xproyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "DVARPROC" & VARYPROC, xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "PVARPROC" & VARYPROC
        ElseIf Not StartRutinas Then
            Call oTrv.Nodes.Add(C_VAR_DSR & k, tvwChild, "DVAR" & VARY, xproyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).aVariables(v).KeyNode = "DVAR" & VARY
        End If
    End If
                    
End Sub

'agrega la funcion al arbol de funciones segun la info del nodo
Private Sub StartChildFuncion(oTrv As TreeView, xproyecto As eProyecto, ByVal NombreFuncion As String, ByVal Icono As Integer, ByVal KeyNode As String)
    
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "FFUN" & Func, NombreFuncion, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "FFUN" & Func
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "BFUN" & Func, NombreFuncion, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "BFUN" & Func
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "CFUN" & Func, NombreFuncion, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "CFUN" & Func
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "KFUN" & Func, NombreFuncion, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "KFUN" & Func
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "PFUN" & Func, NombreFuncion, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "PFUN" & Func
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "DFUN" & Func, NombreFuncion, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "DFUN" & Func
    End If
        
End Sub

'agrega la sub segun el tipo definido
Private Sub StartChildSub(oTrv As TreeView, xproyecto As eProyecto, ByVal NombreSub As String, ByVal Icono As Integer, ByVal KeyNode As String)

    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "FPROC" & PROC, NombreSub, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "FPROC" & PROC
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "BPROC" & PROC, NombreSub, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "BPROC" & PROC
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "CPROC" & PROC, NombreSub, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "CPROC" & PROC
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "KPROC" & PROC, NombreSub, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "KPROC" & PROC
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "PPROC" & PROC, NombreSub, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "PPROC" & PROC
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(KeyNode, tvwChild, "DPROC" & PROC, NombreSub, Icono, Icono)
        xproyecto.aArchivos(k).aRutinas(r).KeyNode = "DPROC" & PROC
    End If
    
End Sub

Private Sub StartChildTipos(oTrv As TreeView, xproyecto As eProyecto, ByVal NombreTipo As String)

    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_TIPOS_FRM & k, tvwChild, "FTIP" & TYPO, NombreTipo, C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).aTipos(t).KeyNode = "FTIP" & TYPO
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_TIPOS_BAS & k, tvwChild, "BTIP" & TYPO, NombreTipo, C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).aTipos(t).KeyNode = "BTIP" & TYPO
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_TIPOS_CLS & k, tvwChild, "CTIP" & TYPO, NombreTipo, C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).aTipos(t).KeyNode = "CTIP" & TYPO
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_TIPOS_CTL & k, tvwChild, "KTIP" & TYPO, NombreTipo, C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).aTipos(t).KeyNode = "KTIP" & TYPO
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_TIPOS_PAG & k, tvwChild, "PTIP" & TYPO, NombreTipo, C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).aTipos(t).KeyNode = "PTIP" & TYPO
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_TIPOS_DSR & k, tvwChild, "DTIP" & TYPO, NombreTipo, C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).aTipos(t).KeyNode = "DTIP" & TYPO
    End If
        
End Sub

'comenzar en el arbol de constantes
Private Sub StartConstantes(oTrv As TreeView, xproyecto As eProyecto)

    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_KEY_FRM & k, tvwChild, C_CONS_FRM & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        xproyecto.aArchivos(k).KeyNodeCte = C_CONS_FRM & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_KEY_BAS & k, tvwChild, C_CONS_BAS & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        xproyecto.aArchivos(k).KeyNodeCte = C_CONS_BAS & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_KEY_CLS & k, tvwChild, C_CONS_CLS & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        xproyecto.aArchivos(k).KeyNodeCte = C_CONS_CLS & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_KEY_CTL & k, tvwChild, C_CONS_CTL & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        xproyecto.aArchivos(k).KeyNodeCte = C_CONS_CTL & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_KEY_PAG & k, tvwChild, C_CONS_PAG & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        xproyecto.aArchivos(k).KeyNodeCte = C_CONS_PAG & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_KEY_DSR & k, tvwChild, C_CONS_DSR & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        xproyecto.aArchivos(k).KeyNodeCte = C_CONS_DSR & k
    End If
    bCon = True
    
End Sub

'colocar icono de variables al arbol del xproyecto
Private Sub StartDim(oTrv As TreeView, xproyecto As eProyecto)

    If Not bVariables And Not StartRutinas Then
        If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call oTrv.Nodes.Add(C_KEY_FRM & k, tvwChild, C_VAR_FRM & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).KeyNodeVar = C_VAR_FRM & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call oTrv.Nodes.Add(C_KEY_BAS & k, tvwChild, C_VAR_BAS & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).KeyNodeVar = C_VAR_BAS & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call oTrv.Nodes.Add(C_KEY_CLS & k, tvwChild, C_VAR_CLS & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).KeyNodeVar = C_VAR_CLS & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call oTrv.Nodes.Add(C_KEY_CTL & k, tvwChild, C_VAR_CTL & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).KeyNodeVar = C_VAR_CTL & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call oTrv.Nodes.Add(C_KEY_PAG & k, tvwChild, C_VAR_PAG & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).KeyNodeVar = C_VAR_PAG & k
        ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            Call oTrv.Nodes.Add(C_KEY_DSR & k, tvwChild, C_VAR_DSR & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            xproyecto.aArchivos(k).KeyNodeVar = C_VAR_DSR & k
        End If
        bVariables = True
    End If
                    
End Sub

'agrega la funcion al arbol
Private Sub StartFuncion(oTrv As TreeView, xproyecto As eProyecto)
    
    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_KEY_FRM & k, tvwChild, C_FUNC_FRM & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        xproyecto.aArchivos(k).KeyNodeFun = C_FUNC_FRM & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_KEY_BAS & k, tvwChild, C_FUNC_BAS & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        xproyecto.aArchivos(k).KeyNodeFun = C_FUNC_BAS & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_KEY_CLS & k, tvwChild, C_FUNC_CLS & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        xproyecto.aArchivos(k).KeyNodeFun = C_FUNC_CLS & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_KEY_CTL & k, tvwChild, C_FUNC_CTL & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        xproyecto.aArchivos(k).KeyNodeFun = C_FUNC_CTL & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_KEY_PAG & k, tvwChild, C_FUNC_PAG & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        xproyecto.aArchivos(k).KeyNodeFun = C_FUNC_PAG & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_KEY_DSR & k, tvwChild, C_FUNC_DSR & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        xproyecto.aArchivos(k).KeyNodeFun = C_FUNC_DSR & k
    End If
    bFun = True
            
End Sub

'comenzar a guardar el codigo de la rutina
Private Sub StartRastreoRutinas(xproyecto As eProyecto)

    If bEndSub Then
        ReDim xproyecto.aArchivos(k).aRutinas(r).aCodigoRutina(0)
        bEndSub = False
    End If
    
End Sub

Private Sub StartSubrutina(oTrv As TreeView, xproyecto As eProyecto)

    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_KEY_FRM & k, tvwChild, C_SUB_FRM & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        xproyecto.aArchivos(k).KeyNodeSub = C_SUB_FRM & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_KEY_BAS & k, tvwChild, C_SUB_BAS & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        xproyecto.aArchivos(k).KeyNodeSub = C_SUB_BAS & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_KEY_CLS & k, tvwChild, C_SUB_CLS & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        xproyecto.aArchivos(k).KeyNodeSub = C_SUB_CLS & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_KEY_CTL & k, tvwChild, C_SUB_CTL & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        xproyecto.aArchivos(k).KeyNodeSub = C_SUB_CTL & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_KEY_PAG & k, tvwChild, C_SUB_PAG & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        xproyecto.aArchivos(k).KeyNodeSub = C_SUB_PAG & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_KEY_DSR & k, tvwChild, C_SUB_DSR & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        xproyecto.aArchivos(k).KeyNodeSub = C_SUB_DSR & k
    End If
    bSub = True
        
End Sub

Private Sub StartTipos(oTrv As TreeView, xproyecto As eProyecto)

    If xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call oTrv.Nodes.Add(C_KEY_FRM & k, tvwChild, C_TIPOS_FRM & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_FRM & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call oTrv.Nodes.Add(C_KEY_BAS & k, tvwChild, C_TIPOS_BAS & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_BAS & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call oTrv.Nodes.Add(C_KEY_CLS & k, tvwChild, C_TIPOS_CLS & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_CLS & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call oTrv.Nodes.Add(C_KEY_CTL & k, tvwChild, C_TIPOS_CTL & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_CTL & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call oTrv.Nodes.Add(C_KEY_PAG & k, tvwChild, C_TIPOS_PAG & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_PAG & k
    ElseIf xproyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call oTrv.Nodes.Add(C_KEY_DSR & k, tvwChild, C_TIPOS_DSR & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        xproyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_DSR & k
    End If
    bTipo = True
            
    
    
End Sub

