VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Project Compare"
   ClientHeight    =   6105
   ClientLeft      =   300
   ClientTop       =   1935
   ClientWidth     =   11280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   752
   WindowState     =   2  'Maximized
   Begin VB.PictureBox SplitterH 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   4245
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  'Custom
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   12
      Tag             =   "0"
      Top             =   2070
      Width           =   1215
   End
   Begin VB.PictureBox picH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   390
      ScaleHeight     =   2220
      ScaleWidth      =   5250
      TabIndex        =   11
      Top             =   3465
      Width           =   5280
      Begin MSComctlLib.ListView lvwProblemas 
         Height          =   1230
         Left            =   300
         TabIndex        =   13
         Top             =   210
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2170
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgProyecto"
         SmallIcons      =   "imgProyecto"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Archivo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ubicación"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Linea"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Dif. Origen->"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "<- Dif.Destino"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Dec. Origen ->"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "<- Dec. Destino"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblDif 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencias entre Proyecto Origen -> Proyecto Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   225
         Left            =   0
         TabIndex        =   14
         Top             =   -15
         Width           =   4665
      End
   End
   Begin MSComctlLib.ProgressBar pgbStatus 
      Height          =   285
      Left            =   1215
      TabIndex        =   10
      Top             =   5490
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox Splitter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   4950
      MouseIcon       =   "frmMain.frx":045C
      MousePointer    =   99  'Custom
      ScaleHeight     =   363
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   9
      Tag             =   "0"
      Top             =   360
      Width           =   45
   End
   Begin VB.PictureBox picDestino 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   1350
      ScaleHeight     =   4665
      ScaleWidth      =   3120
      TabIndex        =   6
      Top             =   450
      Width           =   3150
      Begin MSComctlLib.TreeView treeProyectoD 
         Height          =   3315
         Left            =   30
         TabIndex        =   7
         Top             =   315
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   5847
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imgProyecto"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblDestino 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proyecto Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   15
         TabIndex        =   8
         Top             =   30
         Width           =   3000
      End
   End
   Begin VB.PictureBox picOrigen 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   7320
      ScaleHeight     =   4665
      ScaleWidth      =   3120
      TabIndex        =   3
      Top             =   690
      Width           =   3150
      Begin MSComctlLib.TreeView treeProyectoO 
         Height          =   3315
         Left            =   30
         TabIndex        =   4
         Top             =   1200
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   5847
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imgProyecto"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblOrigen 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proyecto Origen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   30
         TabIndex        =   5
         Top             =   840
         Width           =   3000
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilsIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   38
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdOpenO"
            Object.ToolTipText     =   "Abrir proyecto origen"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdOpenD"
            Object.ToolTipText     =   "Abrir proyecto destino"
            ImageIndex      =   39
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPrint"
            Object.ToolTipText     =   "Imprimir diferencias"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCopiar"
            Object.ToolTipText     =   "Copiar texto al portapapeles"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdBuscar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdComparar"
            Object.ToolTipText     =   "Comparar proyectos segun opciones definidas"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdProyectos"
            Object.ToolTipText     =   "Proyectos"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdArchivos"
            Object.ToolTipText     =   "Archivos"
            ImageIndex      =   30
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdReferencias"
            Object.ToolTipText     =   "Referencias"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdComponentes"
            Object.ToolTipText     =   "Componentes"
            ImageIndex      =   29
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdFormularios"
            Object.ToolTipText     =   "Formularios"
            ImageIndex      =   34
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdModulosBas"
            Object.ToolTipText     =   "Modulos .BAS"
            ImageIndex      =   35
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdModulosCLS"
            Object.ToolTipText     =   "Modulos .CLS"
            ImageIndex      =   36
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdControlesU"
            Object.ToolTipText     =   "Controles de Usuario"
            ImageIndex      =   37
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPaginasP"
            Object.ToolTipText     =   "Paginas de Propiedades"
            ImageIndex      =   38
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSubs"
            Object.ToolTipText     =   "Subs"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdFunciones"
            Object.ToolTipText     =   "Funciones"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdApis"
            Object.ToolTipText     =   "Apis"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPropiedades"
            Object.ToolTipText     =   "Propiedades"
            ImageIndex      =   31
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdVariables"
            Object.ToolTipText     =   "Variables"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdConstantes"
            Object.ToolTipText     =   "Constantes"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdTipos"
            Object.ToolTipText     =   "Tipos"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdEnumeraciones"
            Object.ToolTipText     =   "Enumeraciones"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdArreglos"
            Object.ToolTipText     =   "Arreglos"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdEventos"
            Object.ToolTipText     =   "Eventos"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdControlesFRM"
            Object.ToolTipText     =   "Controles formularios"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdWeb"
            Object.ToolTipText     =   "Ir al sitio web de vbsoftware"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAyuda"
            Object.ToolTipText     =   "Ayuda"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSalir"
            Object.ToolTipText     =   "Salir "
            ImageIndex      =   16
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4725
      Left            =   0
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   705
      Width           =   360
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   5925
      Top             =   5070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05AE
            Key             =   ""
            Object.Tag             =   "Abrir proyecto &origen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
            Object.Tag             =   "&Respaldar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BE6
            Key             =   ""
            Object.Tag             =   "&Imprimir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F02
            Key             =   ""
            Object.Tag             =   "Cortar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":121E
            Key             =   ""
            Object.Tag             =   "&Copiar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153A
            Key             =   ""
            Object.Tag             =   "Pegar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1856
            Key             =   ""
            Object.Tag             =   "&Buscar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21AA
            Key             =   ""
            Object.Tag             =   "Configurar &Opciones a analizar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24C6
            Key             =   ""
            Object.Tag             =   "&Indice"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27E2
            Key             =   ""
            Object.Tag             =   "www"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E1A
            Key             =   ""
            Object.Tag             =   "&Documentar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3136
            Key             =   ""
            Object.Tag             =   "&Comparar"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3452
            Key             =   ""
            Object.Tag             =   "&Salir"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35CA
            Key             =   ""
            Object.Tag             =   "&Proyecto"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3726
            Key             =   ""
            Object.Tag             =   "&Subrutinas"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3882
            Key             =   ""
            Object.Tag             =   "&Funciones"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39DE
            Key             =   ""
            Object.Tag             =   "&Apis"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B3A
            Key             =   ""
            Object.Tag             =   "&Variables"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C96
            Key             =   ""
            Object.Tag             =   "Cons&tantes"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DF2
            Key             =   ""
            Object.Tag             =   "T&ipos"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F4E
            Key             =   ""
            Object.Tag             =   "&Enumeraciones"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40AA
            Key             =   ""
            Object.Tag             =   "Arre&glos"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4206
            Key             =   ""
            Object.Tag             =   "Con&troles"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4362
            Key             =   ""
            Object.Tag             =   "A&brir"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44C2
            Key             =   ""
            Object.Tag             =   "&Referencias"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":461E
            Key             =   ""
            Object.Tag             =   "&Componentes"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":477A
            Key             =   ""
            Object.Tag             =   "Arc&hivos"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48D6
            Key             =   ""
            Object.Tag             =   "Propie&dades"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A30
            Key             =   ""
            Object.Tag             =   "E&ventos"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D66
            Key             =   ""
            Object.Tag             =   "F&ormularios"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4ECE
            Key             =   ""
            Object.Tag             =   "Modulos .&BAS"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5036
            Key             =   ""
            Object.Tag             =   "Modulos .C&LS"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":519E
            Key             =   ""
            Object.Tag             =   "Controles de &Usuario"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5306
            Key             =   ""
            Object.Tag             =   "Pa&ginas de Propiedades"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":546E
            Key             =   ""
            Object.Tag             =   "Abrir proyecto &destino"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgProyecto 
      Left            =   6570
      Top             =   5100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":599A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":613A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6322
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":650A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":707A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7262
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":744A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7632
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":781A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":81A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":838A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8572
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":86CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":88B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9056
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":923E
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9426
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9582
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   5790
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivo_AbrirO 
         Caption         =   "|Seleccionar proyecto origen a comparar|Abrir proyecto &origen"
      End
      Begin VB.Menu mnuArchivo_AbrirD 
         Caption         =   "|Seleccionar proyecto destino a comparar|Abrir proyecto &destino"
      End
      Begin VB.Menu mnuArchivo_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivo_ConfImpresora 
         Caption         =   "|Configurar impresora|&Configurar impresora"
      End
      Begin VB.Menu mnuArchivo_Impresora 
         Caption         =   "|Imprimir diferencias en impresora|&Imprimir"
      End
      Begin VB.Menu mnuArchivo_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivo_Salir 
         Caption         =   "|Salir de la aplicacion|&Salir"
      End
   End
   Begin VB.Menu mnuComparar 
      Caption         =   "&Comparar"
      Enabled         =   0   'False
      Begin VB.Menu mnuComparar_Comparar 
         Caption         =   "|Comparar proyectos seleccionados|&Comparar proyectos"
      End
      Begin VB.Menu mnuComparar_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComparar_Proyecto 
         Caption         =   "|Ver diferencias a nivel de proyectos|&Proyecto"
      End
      Begin VB.Menu mnuComparar_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComparar_Archivos 
         Caption         =   "|Ver diferencias entre archivos del proyecto|Arc&hivos"
      End
      Begin VB.Menu mnuComparar_Referencias 
         Caption         =   "|Ver diferencias entre referencias|&Referencias"
      End
      Begin VB.Menu mnuComparar_Componentes 
         Caption         =   "|Ver diferencias entre componentes|&Componentes"
      End
      Begin VB.Menu mnuComparar_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComparar_Formularios 
         Caption         =   "|Ver diferencias entre formularios|F&ormularios"
      End
      Begin VB.Menu mnuComparar_ModulosBAS 
         Caption         =   "|Ver diferencias entre módulos .bas|Modulos .&BAS"
      End
      Begin VB.Menu mnuComparar_ModulosCLS 
         Caption         =   "|Ver diferencias entre módulos cls|Modulos .C&LS"
      End
      Begin VB.Menu mnuComparar_ControlesUsuario 
         Caption         =   "|Ver diferencias entre controles de usuarios|Controles de &Usuario"
      End
      Begin VB.Menu mnuComparar_Paginas 
         Caption         =   "|Ver diferencias entre páginas de propiedades|Pa&ginas de Propiedades"
      End
      Begin VB.Menu mnuComparar_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComparar_Procedimientos 
         Caption         =   "|Ver diferencias entre subrutinas|&Subrutinas"
      End
      Begin VB.Menu mnuComparar_Funciones 
         Caption         =   "|Ver diferencias entre funciones|&Funciones"
      End
      Begin VB.Menu mnuComparar_Apis 
         Caption         =   "|Ver diferencias entre apis|&Apis"
      End
      Begin VB.Menu mnuComparar_Propiedades 
         Caption         =   "|Ver diferencias entre propiedades|Propie&dades"
      End
      Begin VB.Menu mnuComparar_sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComparar_Variables 
         Caption         =   "|Ver diferencias entre variables|&Variables"
      End
      Begin VB.Menu mnuComparar_Constantes 
         Caption         =   "|Ver diferencias entre constantes|Cons&tantes"
      End
      Begin VB.Menu mnuComparar_Tipos 
         Caption         =   "|Ver diferencias entre tipos|T&ipos"
      End
      Begin VB.Menu mnuComparar_Enumeraciones 
         Caption         =   "|Ver diferencias entre enumeraciones|&Enumeraciones"
      End
      Begin VB.Menu mnuComparar_Arreglos 
         Caption         =   "|Ver diferencias entre arreglos|Arre&glos"
      End
      Begin VB.Menu mnuComparar_Eventos 
         Caption         =   "|Ver diferencias entre eventos|E&ventos"
      End
      Begin VB.Menu mnuComparar_ControlesForm 
         Caption         =   "|Ver diferencias entre controles de formulario|Controles for&mularios"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuOpciones_Configurar 
         Caption         =   "&Configurar opciones de comparación"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAyuda_Indice 
         Caption         =   "&Indice"
      End
      Begin VB.Menu mnuAyuda_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_AcercaDe 
         Caption         =   "Acerca &de ..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private cc As New GCommonDialog
Private clsXmenu As New CXtremeMenu
Private WithEvents MyHelpCallBack As HelpCallBack
Attribute MyHelpCallBack.VB_VarHelpID = -1
Private Itmx As ListItem

Private Const MIN_VERT_BUFFER As Integer = 20
Private Const MIN_HORZ_BUFFER As Integer = 13
Private Const CURSOR_DEDUCT As Integer = 10
Private Const SPLT_WDTH As Integer = 4
Private Const SPLT_HEIGHT As Integer = 4
Private Const CTRL_OFFSET As Integer = 28
Private fInitiateDrag As Boolean
Private Cargando As Boolean
Private Sub AbreNodo(ByVal OD As String, Nodo As Node, ByVal fAbrir As Boolean)

    On Local Error GoTo Salir
    
    Dim Nodox As Node
    Dim k As Integer
    Dim Nombre As String
    Dim Nombre2 As String
    
    If Not Cargando And treeProyectoO.Nodes.Count > 0 And treeProyectoD.Nodes.Count > 0 Then
        Nombre = Nodo.Text
        If InStr(1, Nombre, "-") Then
            Nombre = Trim$(Left$(Nombre, InStr(1, Nombre, "-") - 1))
        ElseIf InStr(1, Nombre, "(") Then
            Nombre = Trim$(Left$(Nombre, InStr(1, Nombre, "(") - 1))
        End If
        If OD = "D" Then
            For k = 1 To treeProyectoO.Nodes.Count
                Nombre2 = treeProyectoO.Nodes(k).Text
                If InStr(1, Nombre2, "-") Then
                    Nombre2 = Trim$(Left$(Nombre2, InStr(1, Nombre2, "-") - 1))
                ElseIf InStr(1, Nombre2, "(") Then
                    Nombre2 = Trim$(Left$(Nombre2, InStr(1, Nombre2, "(") - 1))
                End If
        
                If Nombre = Nombre2 Then
                    If fAbrir Then
                        If Len(Nodo.Key) > 0 And Nodo.Key = treeProyectoO.Nodes(k).Key Then
                            If Not treeProyectoO.Nodes(Nodo.Key).Expanded Then
                                treeProyectoO.Nodes(Nodo.Key).Expanded = fAbrir
                            End If
                        Else
                            Exit For
                        End If
                    Else
                        If Len(Nodo.Key) > 0 And Nodo.Key = treeProyectoO.Nodes(k).Key Then
                            If treeProyectoO.Nodes(Nodo.Key).Expanded Then
                                treeProyectoO.Nodes(Nodo.Key).Expanded = fAbrir
                            End If
                        Else
                            Exit For
                        End If
                    End If
                    Exit For
                End If
            Next k
        Else
            For k = 1 To treeProyectoD.Nodes.Count
                Nombre2 = treeProyectoD.Nodes(k).Text
                If InStr(1, Nombre2, "-") Then
                    Nombre2 = Trim$(Left$(Nombre2, InStr(1, Nombre2, "-") - 1))
                ElseIf InStr(1, Nombre2, "(") Then
                    Nombre2 = Trim$(Left$(Nombre2, InStr(1, Nombre2, "(") - 1))
                End If
                If Nombre = Nombre2 Then
                    If fAbrir Then
                        If Len(Nodo.Key) > 0 And Nodo.Key = treeProyectoD.Nodes(k).Key Then
                            If Not treeProyectoD.Nodes(Nodo.Key).Expanded Then
                                treeProyectoD.Nodes(Nodo.Key).Expanded = fAbrir
                            End If
                        Else
                            Exit For
                        End If
                    Else
                        If Len(Nodo.Key) > 0 And Nodo.Key = treeProyectoO.Nodes(k).Key Then
                            If treeProyectoD.Nodes(Nodo.Key).Expanded Then
                                treeProyectoD.Nodes(Nodo.Key).Expanded = fAbrir
                            End If
                        Else
                            Exit For
                        End If
                    End If
                    Exit For
                End If
            Next k
        End If
    End If
    
    Exit Sub
    
Salir:
    Err = 0
    
End Sub

Private Sub habilitaComparar()

    Dim k As Integer
    If Len(ProyectoO.Archivo) > 0 And Len(ProyectoD.Archivo) > 0 Then
        Call HabilitaTB(False)
        
        tlbMain.Buttons(7).Enabled = True
        tlbMain.Buttons(8).Enabled = True
        tlbMain.Buttons(9).Enabled = True
        
        For k = 11 To tlbMain.Buttons.Count
            tlbMain.Buttons(k).Enabled = False
        Next k
        
        tlbMain.Buttons(tlbMain.Buttons.Count - 3).Enabled = True
        tlbMain.Buttons(tlbMain.Buttons.Count - 2).Enabled = True
        tlbMain.Buttons(tlbMain.Buttons.Count).Enabled = True
    Else
        Call HabilitaTB(False)
    End If
    
End Sub

Private Sub HabilitaTB(ByVal Estado As Boolean)

    Dim k As Integer
        
    For k = 5 To tlbMain.Buttons.Count - 1
        tlbMain.Buttons(k).Enabled = Estado
    Next k
    
End Sub




'imprimir diferencias
Private Function ImprimirDiferencias() As Boolean

    On Local Error GoTo ErrorImprimirDiferencias
    
    Dim ret As Boolean
    Dim k As Integer
    Dim nFreeFile As Integer
    Dim Fuente As String
    Dim Itmx As ListItem
    
    ret = True
    
    Call Hourglass(hWnd, True)
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open App.Path & "\diferencias.htm" For Output As #nFreeFile
        'crear cabezera archivo .html
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Diferencias</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Archivo Origen    : " & ProyectoO.Archivo & "</p>"
        Print #nFreeFile, "<p>Archivo Destino   : " & ProyectoD.Archivo & "</p>"
        Print #nFreeFile, "<p>Total Diferencias : " & lvwProblemas.ListItems.Count & "</p>"
        Print #nFreeFile, "</font>" & "</b>"
        
        'cabezera del informe
        Print #nFreeFile, "<table width='90%' border='1' bordercolor='#FFFFFF'>"
        Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
        Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
        Print #nFreeFile, "<td width='10%'>" & Fuente & "Archivo</font></td>"
        Print #nFreeFile, "<td width='10%'>" & Fuente & "Ubicacion</font></td>"
        Print #nFreeFile, "<td width='05%'>" & Fuente & "Linea</font></td>"
        Print #nFreeFile, "<td width='15%'>" & Fuente & "Dif. Origen</font></td>"
        Print #nFreeFile, "<td width='15%'>" & Fuente & "Dif. Destino</font></td>"
        Print #nFreeFile, "<td width='15%'>" & Fuente & "Dec. Origen</font></td>"
        Print #nFreeFile, "<td width='15%'>" & Fuente & "Dec. Destino</font></td>"
        Print #nFreeFile, "</tr>"
            
        'ciclar x las diferencias
        For k = 1 To lvwProblemas.ListItems.Count
            Set Itmx = lvwProblemas.ListItems(k)
            
            'llenar filas
            Print #nFreeFile, "<tr bordercolor='#000000'>"
            Print #nFreeFile, "<td width='05%'><b>" & Fuente & k & "</font></b></td>"
            Print #nFreeFile, "<td width='10%'>" & Fuente & IIf(Len(Itmx.SubItems(1)) > 0, Itmx.SubItems(1), " ") & "</font></td>"
            Print #nFreeFile, "<td width='10%'>" & Fuente & IIf(Len(Itmx.SubItems(2)) > 0, Itmx.SubItems(2), " ") & "</font></td>"
            Print #nFreeFile, "<td width='05%'>" & Fuente & IIf(Len(Itmx.SubItems(3)) > 0, Itmx.SubItems(3), " ") & "</font></td>"
            Print #nFreeFile, "<td width='15%'>" & Fuente & IIf(Len(Itmx.SubItems(4)) > 0, Itmx.SubItems(4), " ") & "</font></td>"
            Print #nFreeFile, "<td width='15%'>" & Fuente & IIf(Len(Itmx.SubItems(5)) > 0, Itmx.SubItems(5), " ") & "</font></td>"
            Print #nFreeFile, "<td width='15%'>" & Fuente & IIf(Len(Itmx.SubItems(6)) > 0, Itmx.SubItems(6), " ") & "</font></td>"
            Print #nFreeFile, "<td width='15%'>" & Fuente & IIf(Len(Itmx.SubItems(7)) > 0, Itmx.SubItems(7), " ") & "</font></td>"
            Print #nFreeFile, "</tr>"
        Next k
        Print #nFreeFile, "</table>"
        Print #nFreeFile, "<br>"
        
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
    
    GoTo SalirImprimirDiferencias
    
ErrorImprimirDiferencias:
    ret = False
    MsgBox "ImprimirDiferencias : " & Err & " " & Error$, vbCritical
    Resume SalirImprimirDiferencias
    
SalirImprimirDiferencias:
    ImprimirDiferencias = ret
    Err = 0
    Call Hourglass(hWnd, False)
    
End Function


Private Sub Form_Load()

    Set MyHelpCallBack = New HelpCallBack
    
    Call clsXmenu.Install(hWnd, MyHelpCallBack, Me.ilsIcons)
    Call clsXmenu.FontName(hWnd, "Tahoma")
    
    lblDif.Tag = lblDif.Caption
    
    Splitter.Move ScaleWidth \ 3, CTRL_OFFSET + 2, SPLT_WDTH, (ScaleHeight - (CTRL_OFFSET * 2)) - 4
    SplitterH.Move picMain.Width + 1, ScaleHeight, ScaleWidth - picMain.Width
    
    glbSeComparo = False
    
    Call HabilitaTB(False)
    Call Form_Resize
    SplitterH.Move picMain.Width + 1, ScaleHeight, ScaleWidth - picMain.Width
    Call CargaOpcionesDeComparacion
        
    ReDim arr_diferencias(0)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim Msg As String
    
    Msg = "Confirma salir de " & App.Title
    
    If Confirma(Msg) = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    If UnloadMode = 1 Then
        Call ClearTreeView(treeProyectoO.hWnd)
        Call ClearTreeView(treeProyectoD.hWnd)
    End If
    
End Sub

Private Sub Form_Resize()

    On Local Error Resume Next
    
    Dim k As Integer
    Dim FrameWidth As Integer
    
    If WindowState <> vbMinimized Then
        DoEvents
        ' maximized, lock update to avoid nasty window flashing
        If WindowState = vbMaximized Then Call LockWindowUpdate(hWnd)
        
        Call Hourglass(hWnd, True)

        ' handle minimum height. if you were to remove the
        ' controlbox you would need to handle minimum width also
        If Height < 3500 Then Height = 3500
        If Width < 3500 Then Width = 3500
        
        picMain.Left = 0
        picMain.Top = tlbMain.Height + 1
        picMain.Height = ScaleHeight - tlbMain.Height - stbMain.Height
                
        ' the width of the window frame
        FrameWidth = ((Width \ Screen.TwipsPerPixelX) - ScaleWidth) \ 2
    
        ' handle a form resize that hides the vertical splitter
        If ((ScaleWidth - CTRL_OFFSET) - (Splitter.Left + Splitter.Width)) < 12 Then
            Splitter.Left = ScaleWidth - ((CTRL_OFFSET * 4) + (FrameWidth * 2))
        End If
                    
        SplitterH.Left = picMain.Width
        SplitterH.Width = ScaleWidth - SplitterH.Left
        
        'ventana de problemas de comparacion
        picH.Left = SplitterH.Left
        picH.Width = SplitterH.Width
        picH.Top = SplitterH.Top + SplitterH.Height
        picH.Height = ScaleHeight - SplitterH.Top + SplitterH.Height - stbMain.Height - tlbMain.Height + SplitterH.Height
        
        'listview de problemas de comparacion
        lblDif.Top = 0
        lblDif.Left = 0
        lblDif.Width = picH.Width * Screen.TwipsPerPixelY
        
        lvwProblemas.Top = lblDif.Height + 1
        lvwProblemas.Left = 0
        lvwProblemas.Height = picH.Height * Screen.TwipsPerPixelX - lblDif.Height
        lvwProblemas.Width = picH.Width * Screen.TwipsPerPixelY
        
        'height y width del picture que contiene el treeview proyecto origen
        Dim height_picOri As Integer
        Dim Width_picOri As Integer
                
        height_picOri = ScaleHeight - tlbMain.Height - stbMain.Height - picH.Height - SplitterH.Height - 1
        Width_picOri = Splitter.Left - Splitter.Width - picMain.Width + 3
                
        picOrigen.Move 24, picMain.Top, Width_picOri, height_picOri
        lblOrigen.Move 0, 0, Width_picOri * Screen.TwipsPerPixelY
        treeProyectoO.Move 0, lblOrigen.Height, Width_picOri * Screen.TwipsPerPixelY, height_picOri * Screen.TwipsPerPixelX - 260
        
        'cambiar el tamaño del splitter
        Splitter.Top = 25
        Splitter.Height = picOrigen.Height
                        
        'height y width del picture que contiene el treeview proyecto destino
        Dim height_picDes As Integer
        Dim Width_picDes As Integer
        Dim left_picDes As Integer
                
        height_picDes = ScaleHeight - tlbMain.Height - stbMain.Height - picH.Height - SplitterH.Height - 1
        Width_picDes = ScaleWidth - Splitter.Width - picMain.Width - picOrigen.Width
                
        picDestino.Move Splitter.Left + 5, picMain.Top, Width_picDes - 2, height_picDes
        lblDestino.Move 0, 0, Width_picDes * Screen.TwipsPerPixelY
        treeProyectoD.Move 0, lblDestino.Height, Width_picDes * Screen.TwipsPerPixelY, height_picDes * Screen.TwipsPerPixelX - 260
                
        pgbStatus.Top = ScaleHeight - 10
        pgbStatus.Left = stbMain.Panels(2).Left + 4
        pgbStatus.Height = stbMain.Height - 5
        pgbStatus.Width = stbMain.Panels(2).Width - 7
        pgbStatus.ZOrder 0

        With mGradient
            .Angle = 90 '.Angle
            .Color1 = 16744448
            .Color2 = 0
            .Draw picMain
        End With
            
        Call FontStuff(picMain, App.Title & " Beta Versión : " & App.Major & "." & App.Minor & "." & App.Revision)
        
        picMain.Refresh
        
        'Call Hourglass(hWnd, False)
        
        If Splitter.Left < 24 Then
            Splitter.Left = 200
            Call Form_Resize
        End If
                        
        If SplitterH.Top < 20 Then
            SplitterH.Top = 300
            Call Form_Resize
        ElseIf SplitterH.Top > 515 Then
            SplitterH.Top = 400
            Call Form_Resize
        End If
        
        Splitter.ZOrder 0
        SplitterH.ZOrder 0
        
        ' if it's locked unlock the window
        If WindowState = vbMaximized Then Call LockWindowUpdate(0&)
        Call Hourglass(hWnd, False)
        DoEvents
    End If
    
    Err = 0
    
End Sub


Private Sub mnuArchivo_AbrirD_Click()

    Dim Archivo As String
    Dim Glosa As String
    
    If gsLastPath = "" Then gsLastPath = App.Path

    Glosa = "Visual Basic 3.0 (*.MAK)|*.MAK|"
    Glosa = Glosa & "Visual Basic 4,5,6 (*.VBP)|*.VBP|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If Not (cc.VBGetOpenFileName(Archivo, , , , , , Glosa, , gsLastPath, "Abrir proyecto Visual Basic...", "VBP", Me.hWnd)) Then
       Exit Sub
    End If
   
    If Archivo = "" Then Exit Sub
    
    ProyectoD.Analizado = False
    
    Cargando = True
    
    Call AnalizaProyectoVB(Archivo, 2)
    
    Call habilitaComparar
    
    Cargando = False
    
End Sub

Private Sub mnuArchivo_AbrirO_Click()
    
    Dim Archivo As String
    Dim Glosa As String
    
    If gsLastPath = "" Then gsLastPath = App.Path

    Glosa = "Visual Basic 3.0 (*.MAK)|*.MAK|"
    Glosa = Glosa & "Visual Basic 4,5,6 (*.VBP)|*.VBP|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If Not (cc.VBGetOpenFileName(Archivo, , , , , , Glosa, , gsLastPath, "Abrir proyecto Visual Basic...", "VBP", Me.hWnd)) Then
       Exit Sub
    End If
   
    If Archivo = "" Then Exit Sub
    
    ProyectoO.Analizado = False
    
    Cargando = True

    Call AnalizaProyectoVB(Archivo, 1)
        
    Call habilitaComparar
    
    Cargando = False
    
End Sub

'abre el proyecto visual basic y lo prepara para el analisis
Public Sub AnalizaProyectoVB(ByVal Archivo As String, ByVal Arbol As Integer)

    Dim k As Integer
    Dim ret As Boolean
    Dim Msg As String
    
    Call Hourglass(hWnd, True)
    
    gsLastPath = PathArchivo(Archivo)
            
    InhabilitaToolbar False
    
    lvwProblemas.ListItems.Clear
    
    If Arbol = 1 Then   'origen
        Call ClearTreeView(treeProyectoO.hWnd)
    Else
        Call ClearTreeView(treeProyectoD.hWnd)
    End If
            
    stbMain.Panels(1).Text = ""
    stbMain.Panels(2).Text = ""
    stbMain.Panels(3).Text = ""
    stbMain.Panels(4).Text = ""
    stbMain.Panels(5).Text = ""
            
    If Arbol = 1 Then   'origen
        ret = CargaProyecto(Archivo, frmMain.treeProyectoO, ProyectoO, TotalesProyectoO)
    Else
        ret = CargaProyecto(Archivo, frmMain.treeProyectoD, ProyectoD, TotalesProyectoD)
    End If
          
    InhabilitaToolbar True
    
End Sub

Private Sub mnuArchivo_Impresora_Click()
    
    Dim Msg As String
    
    If lvwProblemas.ListItems.Count > 0 Then
        Msg = "Confirma imprimir diferencias."
        If Confirma(Msg) = vbYes Then
            If ImprimirDiferencias() Then
                MsgBox "Informe generado con éxito!", vbInformation
                On Local Error Resume Next
                ShellExecute Me.hWnd, vbNullString, App.Path & "\diferencias.htm", vbNullString, App.Path & "\", SW_SHOWMAXIMIZED
                Err = 0
            End If
        End If
    End If
    
End Sub

Private Sub mnuArchivo_Salir_Click()
    Unload Me
End Sub

Private Sub mnuAyuda_AcercaDe_Click()
    frmAcerca.Show vbModal
End Sub

Private Sub mnuComparar_Apis_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Apis")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Archivos_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Archivos")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Arreglos_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Arreglos")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Comparar_Click()

    Dim Msg As String
    
    Msg = "Confirma comparar proyectos."
    
    If Confirma(Msg) = vbYes Then
        Call InhabilitaToolbar(False)
        mnuComparar.Enabled = False
        frmMain.lblDif.Caption = frmMain.lblDif.Tag
        Call Form_Resize
        If CompararProyectosSeleccionados Then
            MsgBox "Comparación de proyectos seleccionados realizada con éxito!", vbInformation
        End If
        Call Form_Resize
        Call InhabilitaToolbar(True)
    End If
    
End Sub

Private Sub mnuComparar_Componentes_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Componentes")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Constantes_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Constantes")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_ControlesForm_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Controles Form")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_ControlesUsuario_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Controles")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Enumeraciones_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Enumeraciones")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Eventos_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Eventos")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Formularios_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Formularios")
    Call InhabilitaToolbar(True)
End Sub


Private Sub mnuComparar_Funciones_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Funciones")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_ModulosBAS_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Modulos .BAS")
    Call InhabilitaToolbar(True)
End Sub


Private Sub mnuComparar_ModulosCLS_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Modulos .CLS")
    Call InhabilitaToolbar(True)
End Sub


Private Sub mnuComparar_Paginas_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Paginas")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Procedimientos_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Procedimientos")
    Call InhabilitaToolbar(True)
End Sub


Private Sub mnuComparar_Propiedades_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Propiedades")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Proyecto_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Proyecto")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Referencias_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Referencias")
    Call InhabilitaToolbar(True)
End Sub


Private Sub mnuComparar_Tipos_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Tipos")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuComparar_Variables_Click()
    Call InhabilitaToolbar(False)
    Call FiltraComparaciones("Variables")
    Call InhabilitaToolbar(True)
End Sub

Private Sub mnuOpciones_Configurar_Click()
    frmOpciones.Show vbModal
End Sub

Private Sub MyHelpCallBack_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
    stbMain.Panels(1).Text = MenuHelp
End Sub

Private Sub Splitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' if the left button is down set the flag
    If Button = 1 Then fInitiateDrag = True
End Sub


Private Sub Splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' if the flag isn't set then the left button wasn't
    ' pressed while the mouse was over one of the splitters
    If fInitiateDrag <> True Then Exit Sub

    ' if the left button is down then we want to move the splitter
    If Button = 1 Then ' if the Tag is false then we need to set
        If Splitter.Tag = False Then ' the color and clip the cursor.
    
            Splitter.BackColor = &H808080 '<- set the "dragging" color here
            Splitter.Tag = True
        End If
    
        Splitter.Left = (Splitter.Left + x) - (SPLT_WDTH \ 3)
    End If
    
End Sub


Private Sub Splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' if the left button is the one being released we need to reset
    ' the color, Tag, flag, cancel ClipCursor and call form_resize
  
    If Button = 1 Then           ' to move the list and text boxes
        Splitter.Tag = False
        fInitiateDrag = False
        'ClipCursor ByVal 0&
        Splitter.BackColor = &H8000000F  '<- set to original color
        Form_Resize
    End If
    
End Sub


Private Sub SplitterH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' if the left button is down set the flag
    If Button = 1 Then fInitiateDrag = True
End Sub

Private Sub SplitterH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' if the flag isn't set then the left button wasn't
    ' pressed while the mouse was over one of the splitters
    If fInitiateDrag <> True Then Exit Sub

    ' if the left button is down then we want to move the splitter
    If Button = 1 Then ' if the Tag is false then we need to set
        If SplitterH.Tag = False Then ' the color and clip the cursor.
    
            Splitter.BackColor = &H808080 '<- set the "dragging" color here
            SplitterH.Tag = True
        End If
    
        SplitterH.Top = (SplitterH.Top + y) - (SPLT_WDTH \ 3)
    End If
    
End Sub

Private Sub SplitterH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' if the left button is the one being released we need to reset
    ' the color, Tag, flag, cancel ClipCursor and call form_resize
  
    If Button = 1 Then           ' to move the list and text boxes
        SplitterH.Tag = False
        fInitiateDrag = False
        Splitter.BackColor = &H8000000F  '<- set to original color
        Form_Resize
    End If
    
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "cmdOpenO"
            mnuArchivo_AbrirO_Click
        Case "cmdOpenD"
            mnuArchivo_AbrirD_Click
        Case "cmdPrint"
            mnuArchivo_Impresora_Click
        Case "cmdBuscar"
            frmBuscar.Show vbModal
        Case "cmdComparar"
            mnuComparar_Comparar_Click
        Case "cmdProyectos"
            mnuComparar_Proyecto_Click
        Case "cmdArchivos"
            mnuComparar_Archivos_Click
        Case "cmdReferencias"
            mnuComparar_Referencias_Click
        Case "cmdComponentes"
            mnuComparar_Componentes_Click
        Case "cmdFormularios"
            mnuComparar_Formularios_Click
        Case "cmdModulosBas"
            mnuComparar_ModulosBAS_Click
        Case "cmdModulosCLS"
            mnuComparar_ModulosCLS_Click
        Case "cmdControlesU"
            mnuComparar_ControlesUsuario_Click
        Case "cmdPaginasP"
            mnuComparar_Paginas_Click
        Case "cmdSubs"
            mnuComparar_Procedimientos_Click
        Case "cmdFunciones"
            mnuComparar_Funciones_Click
        Case "cmdApis"
            mnuComparar_Apis_Click
        Case "cmdPropiedades"
            mnuComparar_Propiedades_Click
        Case "cmdVariables"
            mnuComparar_Variables_Click
        Case "cmdConstantes"
            mnuComparar_Constantes_Click
        Case "cmdTipos"
            mnuComparar_Tipos_Click
        Case "cmdEnumeraciones"
            mnuComparar_Enumeraciones_Click
        Case "cmdArreglos"
            mnuComparar_Arreglos_Click
        Case "cmdEventos"
            mnuComparar_Eventos_Click
        Case "cmdControlesFRM"
            mnuComparar_ControlesForm_Click
        Case "cmdWeb"
            pShell C_WEB_PAGE, hWnd
        Case "cmdSalir"
            mnuArchivo_Salir_Click
    End Select
    
End Sub
Private Sub InhabilitaToolbar(ByVal Estado As Boolean)

    Dim k As Integer
            
    mnuArchivo.Enabled = Estado
    mnuComparar.Enabled = glbSeComparo
    mnuOpciones.Enabled = Estado
    mnuAyuda.Enabled = Estado
    
    'esperar a que llegue algun registro
    For k = 1 To tlbMain.Buttons.Count
        tlbMain.Buttons(k).Enabled = Estado
    Next k
    
End Sub


Private Sub treeProyectoD_Collapse(ByVal Node As MSComctlLib.Node)

    Select Case Node.Text
        Case "Referencias", "Componentes", "Formularios", "Módulos", "Módulos de Clase"
            Node.SelectedImage = C_ICONO_CLOSE
            Node.Image = C_ICONO_CLOSE
            Call AbreNodo("D", Node, False)
        Case "Controles de Usuario", "Páginas de Propiedades", "Documentos Relacionados", "Diseñadores"
            Node.SelectedImage = C_ICONO_CLOSE
            Node.Image = C_ICONO_CLOSE
            Call AbreNodo("D", Node, False)
        Case Else
            Call AbreNodo("D", Node, False)
    End Select
    
End Sub


Private Sub treeProyectoD_Expand(ByVal Node As MSComctlLib.Node)

    Select Case Node.Text
        Case "Referencias", "Componentes", "Formularios", "Módulos", "Módulos de Clase"
            Node.Image = C_ICONO_OPEN
            Node.SelectedImage = C_ICONO_OPEN
            Call AbreNodo("D", Node, True)
        Case "Controles de Usuario", "Páginas de Propiedades", "Documentos Relacionados", "Diseñadores"
            Node.Image = C_ICONO_OPEN
            Node.SelectedImage = C_ICONO_OPEN
            Call AbreNodo("D", Node, True)
        Case Else
            Call AbreNodo("D", Node, True)
    End Select
    
End Sub


Private Sub treeProyectoD_NodeClick(ByVal Node As MSComctlLib.Node)
    Call AbreNodo("D", Node, True)
End Sub

Private Sub treeProyectoO_Collapse(ByVal Node As MSComctlLib.Node)

     Select Case Node.Text
        Case "Referencias", "Componentes", "Formularios", "Módulos", "Módulos de Clase"
            Node.SelectedImage = C_ICONO_CLOSE
            Node.Image = C_ICONO_CLOSE
            Call AbreNodo("O", Node, False)
        Case "Controles de Usuario", "Páginas de Propiedades", "Documentos Relacionados", "Diseñadores"
            Node.SelectedImage = C_ICONO_CLOSE
            Node.Image = C_ICONO_CLOSE
            Call AbreNodo("O", Node, False)
        Case Else
            Call AbreNodo("O", Node, False)
    End Select
    
End Sub


Private Sub treeProyectoO_Expand(ByVal Node As MSComctlLib.Node)

    Select Case Node.Text
        Case "Referencias", "Componentes", "Formularios", "Módulos", "Módulos de Clase"
            Node.Image = C_ICONO_OPEN
            Node.SelectedImage = C_ICONO_OPEN
            Call AbreNodo("O", Node, True)
        Case "Controles de Usuario", "Páginas de Propiedades", "Documentos Relacionados", "Diseñadores"
            Node.Image = C_ICONO_OPEN
            Node.SelectedImage = C_ICONO_OPEN
            Call AbreNodo("O", Node, True)
        Case Else
            Call AbreNodo("O", Node, True)
    End Select
    
End Sub


Private Sub treeProyectoO_NodeClick(ByVal Node As MSComctlLib.Node)

    Call AbreNodo("O", Node, True)
    
End Sub

