VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Comparación"
   ClientHeight    =   3750
   ClientLeft      =   1305
   ClientTop       =   1755
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Opciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOpci 
      Caption         =   "Opciones Generales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Index           =   3
      Left            =   900
      TabIndex        =   28
      Top             =   4200
      Width           =   4140
      Begin VB.Frame fraAyuda 
         Caption         =   "Ayuda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   135
         TabIndex        =   30
         Top             =   555
         Width           =   3915
         Begin VB.Label lblAyuda 
            Caption         =   $"Opciones.frx":030A
            Height          =   1245
            Left            =   120
            TabIndex        =   31
            Top             =   225
            Width           =   3705
         End
      End
      Begin VB.CheckBox chkGen 
         Caption         =   "Comparar linea a linea"
         Height          =   225
         Left            =   135
         TabIndex        =   29
         Top             =   285
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   4965
      TabIndex        =   6
      Top             =   810
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   4965
      TabIndex        =   5
      Top             =   330
      Width           =   1215
   End
   Begin VB.Frame fraOpci 
      Caption         =   "Opciones de codigo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Index           =   2
      Left            =   1650
      TabIndex        =   4
      Top             =   4350
      Width           =   3225
      Begin VB.CheckBox chkCodigo 
         Caption         =   "Arreglos"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   17
         Top             =   1515
         Width           =   1005
      End
      Begin VB.CheckBox chkCodigo 
         Caption         =   "Eventos"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   16
         Top             =   2715
         Width           =   1050
      End
      Begin VB.CheckBox chkCodigo 
         Caption         =   "Procedimientos"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   15
         Top             =   2475
         Width           =   1440
      End
      Begin VB.CheckBox chkCodigo 
         Caption         =   "Funciones"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   14
         Top             =   2235
         Width           =   1110
      End
      Begin VB.CheckBox chkCodigo 
         Caption         =   "Propiedades"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   13
         Top             =   1995
         Width           =   1320
      End
      Begin VB.CheckBox chkCodigo 
         Caption         =   "Apis"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   12
         Top             =   1755
         Width           =   705
      End
      Begin VB.CheckBox chkCodigo 
         Caption         =   "Tipos"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   11
         Top             =   1290
         Width           =   780
      End
      Begin VB.CheckBox chkCodigo 
         Caption         =   "Enumeraciones"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   1050
         Width           =   1515
      End
      Begin VB.CheckBox chkCodigo 
         Caption         =   "Constantes"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   810
         Width           =   1200
      End
      Begin VB.CheckBox chkCodigo 
         Caption         =   "Variables"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   570
         Width           =   1170
      End
      Begin VB.CheckBox chkCodigo 
         Caption         =   "General del archivo"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   330
         Width           =   1830
      End
   End
   Begin VB.Frame fraOpci 
      Caption         =   "Opciones de archivos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3030
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   4560
      Width           =   3225
      Begin VB.CheckBox chkArchivos 
         Caption         =   "Diseñadores"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   23
         Top             =   1575
         Width           =   1365
      End
      Begin VB.CheckBox chkArchivos 
         Caption         =   "Páginas de Propiedades"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   22
         Top             =   1320
         Width           =   2130
      End
      Begin VB.CheckBox chkArchivos 
         Caption         =   "Controles de Usuario"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   21
         Top             =   1050
         Width           =   2070
      End
      Begin VB.CheckBox chkArchivos 
         Caption         =   "Módulos .CLS"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   20
         Top             =   795
         Width           =   1275
      End
      Begin VB.CheckBox chkArchivos 
         Caption         =   "Módulos .BAS"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   19
         Top             =   540
         Width           =   1350
      End
      Begin VB.CheckBox chkArchivos 
         Caption         =   "Formularios"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   285
         Width           =   1200
      End
   End
   Begin VB.Frame fraOpci 
      Caption         =   "Opciones de proyecto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Index           =   0
      Left            =   435
      TabIndex        =   2
      Top             =   375
      Width           =   4380
      Begin VB.CheckBox chkProyecto 
         Caption         =   "Archivos"
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   27
         Top             =   1080
         Width           =   1185
      End
      Begin VB.CheckBox chkProyecto 
         Caption         =   "Referencias"
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   26
         Top             =   810
         Width           =   1335
      End
      Begin VB.CheckBox chkProyecto 
         Caption         =   "Componentes"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   25
         Top             =   525
         Width           =   1440
      End
      Begin VB.CheckBox chkProyecto 
         Caption         =   "Información de proyecto"
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   24
         Top             =   270
         Width           =   2175
      End
   End
   Begin MSComctlLib.TabStrip tabOpc 
      Height          =   3735
      Left            =   375
      TabIndex        =   1
      Top             =   0
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   6588
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Proyecto"
            Object.ToolTipText     =   "Opciones de comparacion de proyectos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Archivos"
            Object.ToolTipText     =   "Opciones de comparacion de archivos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Codigo"
            Object.ToolTipText     =   "Opciones de comparacion de codigo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Miscélaneas"
            Object.ToolTipText     =   "Opciones anexas"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
'graba las opciones de comparacion
Private Sub GrabarOpciones()

    Dim k As Integer
    
    For k = 0 To 3
        Call GrabaIni(C_INI, "Proyecto", "Valor" & k + 1, chkProyecto(k).Value)
        arr_ComProyecto(k + 1) = chkProyecto(k).Value
    Next k
    
    For k = 0 To 5
        Call GrabaIni(C_INI, "Archivos", "Valor" & k + 1, chkArchivos(k).Value)
        arr_ComArchivos(k + 1) = chkArchivos(k).Value
    Next k
    
    For k = 0 To 10
        Call GrabaIni(C_INI, "Codigo", "Valor" & k + 1, chkCodigo(k).Value)
        arr_ComCodigo(k + 1) = chkCodigo(k).Value
    Next k
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call GrabarOpciones
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Dim k As Integer
    
    Call Hourglass(hWnd, True)
    
    CenterWindow hWnd
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, Me.Caption)
    
    picDraw.Refresh
    
    fraOpci(1).Left = fraOpci(0).Left
    fraOpci(1).Height = fraOpci(0).Height
    fraOpci(1).Width = fraOpci(0).Width
    fraOpci(1).Top = fraOpci(0).Top
    fraOpci(1).Visible = False
    
    fraOpci(2).Left = fraOpci(0).Left
    fraOpci(2).Height = fraOpci(0).Height
    fraOpci(2).Width = fraOpci(0).Width
    fraOpci(2).Top = fraOpci(0).Top
    fraOpci(2).Visible = False
    
    fraOpci(3).Left = fraOpci(0).Left
    fraOpci(3).Height = fraOpci(0).Height
    fraOpci(3).Width = fraOpci(0).Width
    fraOpci(3).Top = fraOpci(0).Top
    fraOpci(3).Visible = False
    
    fraOpci(0).ZOrder 0
    
    For k = 1 To UBound(arr_ComProyecto)
        chkProyecto(k - 1).Value = arr_ComProyecto(k)
    Next k
    
    For k = 1 To UBound(arr_ComArchivos)
        chkArchivos(k - 1).Value = arr_ComArchivos(k)
    Next k
    
    For k = 1 To UBound(arr_ComCodigo)
        chkCodigo(k - 1).Value = arr_ComCodigo(k)
    Next k
    
    chkGen.Value = glbOpcionComparar
    
    Call Hourglass(hWnd, False)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmOpciones = Nothing
End Sub


Private Sub tabOpc_Click()

    If tabOpc.SelectedItem.Index = 1 Then
        fraOpci(0).ZOrder 0
        fraOpci(0).Visible = True
    ElseIf tabOpc.SelectedItem.Index = 2 Then
        fraOpci(1).ZOrder 0
        fraOpci(1).Visible = True
    ElseIf tabOpc.SelectedItem.Index = 3 Then
        fraOpci(2).ZOrder 0
        fraOpci(2).Visible = True
    Else
        fraOpci(3).ZOrder 0
        fraOpci(3).Visible = True
    End If
    
End Sub


