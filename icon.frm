VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmNotfIc 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualização Remota Rm -> RodBel, Versão 14/05/2020 - v2.02"
   ClientHeight    =   6375
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13965
   Icon            =   "icon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   13965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame12 
      Caption         =   "Ajuste das Batidas Ponto"
      DragIcon        =   "icon.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1245
      Left            =   7440
      TabIndex        =   78
      Top             =   9390
      Width           =   2595
      Begin MSComCtl2.DTPicker txt_time_Hora_Bat_Ponto 
         Height          =   315
         Left            =   510
         TabIndex        =   79
         Top             =   630
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139132930
         UpDown          =   -1  'True
         CurrentDate     =   39716.875
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Na Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   630
         TabIndex        =   80
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Afastados/Vale Transp-GED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1245
      Left            =   840
      TabIndex        =   73
      ToolTipText     =   "Ajuste os horários a que horas o sistema vai fazer as novas Atualizações dos funcionários afastados"
      Top             =   9450
      Width           =   2625
      Begin MSComCtl2.DTPicker txt_time_Func_Afast_Ini 
         Height          =   315
         Left            =   1140
         TabIndex        =   74
         Top             =   390
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   32768
         CalendarTitleForeColor=   32768
         Format          =   139132930
         UpDown          =   -1  'True
         CurrentDate     =   39716.5208333333
      End
      Begin MSComCtl2.DTPicker txt_time_Func_Afast_Fim 
         Height          =   315
         Left            =   1140
         TabIndex        =   75
         Top             =   780
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   32768
         Format          =   139132930
         UpDown          =   -1  'True
         CurrentDate     =   39716.9791666667
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "1a Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   90
         TabIndex        =   77
         Top             =   390
         Width           =   960
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "2a Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   60
         TabIndex        =   76
         Top             =   1020
         Width           =   960
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "H.Extras Realizadas p/ Ged"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1245
      Left            =   11490
      TabIndex        =   70
      Top             =   9180
      Width           =   2595
      Begin MSComCtl2.DTPicker txt_time_Hora_Real_Ged 
         Height          =   315
         Left            =   510
         TabIndex        =   71
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139395074
         UpDown          =   -1  'True
         CurrentDate     =   39716.0833333333
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Na Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   690
         TabIndex        =   72
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cad.Função/Seção p/Ged"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1245
      Left            =   10200
      TabIndex        =   67
      Top             =   9270
      Width           =   2595
      Begin MSComCtl2.DTPicker txt_time_Cad_Funcao_Secao_Ged 
         Height          =   315
         Left            =   480
         TabIndex        =   68
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139395074
         UpDown          =   -1  'True
         CurrentDate     =   39716.125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "a Cada Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   420
         TabIndex        =   69
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "Mudar Status NFE"
      DragIcon        =   "icon.frx":074C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1245
      Left            =   2310
      TabIndex        =   64
      Top             =   9450
      Width           =   2595
      Begin MSComCtl2.DTPicker txt_time_Hora_Status_NFE 
         Height          =   315
         Left            =   660
         TabIndex        =   65
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139395074
         UpDown          =   -1  'True
         CurrentDate     =   39716.0034722222
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "A cada Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   630
         TabIndex        =   66
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "He.Extrap./Não autorizadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1245
      Left            =   6150
      TabIndex        =   61
      ToolTipText     =   "Ajuste o minuto e segundo, que o sistema verificará os funcionários desligados. Será feita no minuto/segundo  de cada hora."
      Top             =   9330
      Width           =   2595
      Begin MSComCtl2.DTPicker txt_time_Func_HE_Ext_Nautor 
         Height          =   315
         Left            =   540
         TabIndex        =   62
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139395074
         UpDown          =   -1  'True
         CurrentDate     =   39716.2083333333
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Três D.Úteis e Na Hora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   90
         TabIndex        =   63
         Top             =   270
         Width           =   2445
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "He.Extrap./Não autorizadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1245
      Left            =   4650
      TabIndex        =   55
      ToolTipText     =   "Ajuste o minuto e segundo, que o sistema verificará os funcionários desligados. Será feita no minuto/segundo  de cada hora."
      Top             =   9390
      Width           =   2595
      Begin VB.TextBox txt_dia_email 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         MaxLength       =   2
         TabIndex        =   56
         Text            =   "18"
         Top             =   510
         Width           =   345
      End
      Begin MSComCtl2.DTPicker txt_time_Func_HE_Ext_Nautor_RH 
         Height          =   315
         Left            =   720
         TabIndex        =   57
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139395074
         UpDown          =   -1  'True
         CurrentDate     =   39716.2083333333
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Email p/ Setor 651"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   210
         TabIndex        =   60
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   120
         TabIndex        =   59
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   120
         TabIndex        =   58
         Top             =   870
         Width           =   585
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Mudança de Turno"
      DragIcon        =   "icon.frx":0A56
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1245
      Left            =   300
      TabIndex        =   51
      Top             =   10050
      Width           =   2595
      Begin MSComCtl2.DTPicker txt_time_Func_Muda_Turno 
         Height          =   315
         Left            =   630
         TabIndex        =   52
         Top             =   630
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139395074
         UpDown          =   -1  'True
         CurrentDate     =   39716.0416666667
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "a cada Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   240
         Left            =   630
         TabIndex        =   53
         Top             =   270
         Width           =   1425
      End
   End
   Begin VB.Frame FrmAcesso 
      BackColor       =   &H00008000&
      Caption         =   "      Digite a senha de Acesso     "
      Height          =   975
      Left            =   90
      TabIndex        =   37
      Top             =   6390
      Width           =   2595
      Begin VB.TextBox txt_senha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   11
         PasswordChar    =   "*"
         TabIndex        =   38
         Text            =   "20080101ACS"
         ToolTipText     =   "Digite a senha para obter acesso a tela de manutenção."
         Top             =   330
         Width           =   2085
      End
   End
   Begin MSComctlLib.ProgressBar Pr_Prog 
      Height          =   435
      Left            =   3060
      TabIndex        =   35
      Top             =   5910
      Visible         =   0   'False
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Frame frm_impressao 
      BackColor       =   &H00FF8080&
      Caption         =   "Filtro das informações do LOG"
      Height          =   2535
      Left            =   4680
      TabIndex        =   20
      Top             =   6750
      Visible         =   0   'False
      Width           =   5745
      Begin VB.CommandButton cmd_Impressao 
         Height          =   645
         Left            =   4590
         Picture         =   "icon.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1800
         Width           =   765
      End
      Begin VB.ComboBox CBO_ACOES2 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2190
         TabIndex        =   26
         Text            =   "CBO_ACOES2"
         Top             =   1290
         Width           =   3165
      End
      Begin MSComCtl2.DTPicker DT_Filtro_ini 
         Height          =   285
         Left            =   1170
         TabIndex        =   22
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139460609
         CurrentDate     =   39720
      End
      Begin MSComCtl2.DTPicker DT_Filtro_fim 
         Height          =   285
         Left            =   3990
         TabIndex        =   23
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139460609
         CurrentDate     =   39720
      End
      Begin MSComCtl2.DTPicker DT_Hora_ini 
         Height          =   315
         Left            =   1170
         TabIndex        =   24
         Top             =   720
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139460610
         UpDown          =   -1  'True
         CurrentDate     =   39716.0000115741
      End
      Begin MSComCtl2.DTPicker DT_Hora_fim 
         Height          =   315
         Left            =   3990
         TabIndex        =   25
         Top             =   720
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   32768
         Format          =   139460610
         UpDown          =   -1  'True
         CurrentDate     =   39716.9999884259
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   5670
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   5670
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Atualização de.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1680
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "H.Final.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2970
         TabIndex        =   30
         Top             =   750
         Width           =   870
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "H.Inicial.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   750
         Width           =   975
      End
      Begin VB.Line Line7 
         X1              =   90
         X2              =   5640
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Dt.Final.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3030
         TabIndex        =   28
         Top             =   300
         Width           =   930
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Dt.Inicial.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmd_Log 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Log das Ações realizadas"
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Clique para aparecer/Desaparecer filtro da impressão do LOG."
      Top             =   5910
      Width           =   2865
   End
   Begin VB.Frame frm_principal 
      Caption         =   "Horários Para Começar Procesos de :"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   11
      Top             =   60
      Width           =   13725
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   585
         Left            =   6840
         TabIndex        =   54
         Top             =   3330
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Frame Frame9 
         Caption         =   "Funcionários Atestado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   1245
         Left            =   5504
         TabIndex        =   40
         ToolTipText     =   "Ajuste os horários a que horas o sistema vai fazer as novas Atualizações dos funcionários afastados"
         Top             =   390
         Width           =   2595
         Begin MSComCtl2.DTPicker txt_time_Func_Atest_Ini 
            Height          =   315
            Left            =   1140
            TabIndex        =   41
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   32768
            CalendarTitleForeColor=   32768
            Format          =   139460610
            UpDown          =   -1  'True
            CurrentDate     =   39716.5208333333
         End
         Begin MSComCtl2.DTPicker txt_time_Func_Atest_Fim 
            Height          =   315
            Left            =   1140
            TabIndex        =   42
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   32768
            Format          =   139460610
            UpDown          =   -1  'True
            CurrentDate     =   39716.9791666667
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "1a Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   240
            Left            =   90
            TabIndex        =   44
            Top             =   390
            Width           =   960
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "2a Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   240
            Left            =   90
            TabIndex        =   43
            Top             =   840
            Width           =   960
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Período de pesquisa de acesso aos dados dos funcionários no RM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   5910
         TabIndex        =   32
         Top             =   4350
         Width           =   7455
         Begin VB.TextBox txt_Dias_Antecedencia 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3270
            MaxLength       =   4
            TabIndex        =   5
            Text            =   "05"
            Top             =   330
            Width           =   525
         End
         Begin MSComCtl2.DTPicker DT_Pesquisa 
            Height          =   345
            Left            =   5400
            TabIndex        =   6
            Top             =   330
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   139460609
            CurrentDate     =   39721
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Dt.Pesquisa.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   3930
            TabIndex        =   34
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Dias de antecedência.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   810
            TabIndex        =   33
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.CommandButton CMD_ACAO 
         BackColor       =   &H0080FF80&
         Caption         =   "Confirma Atualização"
         Height          =   825
         Left            =   180
         Picture         =   "icon.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   4800
         Width           =   5265
      End
      Begin VB.ComboBox CBO_ACOES 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4440
         Width           =   5265
      End
      Begin VB.Frame Frame6 
         Caption         =   "Mudança de Setor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   1245
         Left            =   8181
         TabIndex        =   18
         ToolTipText     =   "Ajuste o minuto, que o sistema verificará mudança de Setor. Será feita no minuto/segundos de cada hora."
         Top             =   390
         Width           =   2625
         Begin MSComCtl2.DTPicker txt_time_Func_Muda_Setor 
            Height          =   315
            Left            =   630
            TabIndex        =   4
            Top             =   630
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   139460610
            UpDown          =   -1  'True
            CurrentDate     =   39716.0416666667
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "No minuto da Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C0C0&
            Height          =   240
            Left            =   390
            TabIndex        =   19
            Top             =   300
            Width           =   2070
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Funcionários Desligados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   10890
         TabIndex        =   17
         ToolTipText     =   "Ajuste o minuto e segundo, que o sistema verificará os funcionários desligados. Será feita no minuto/segundo  de cada hora."
         Top             =   390
         Width           =   2595
         Begin MSComCtl2.DTPicker txt_time_Func_Desligados 
            Height          =   315
            Left            =   630
            TabIndex        =   2
            Top             =   630
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   139460610
            UpDown          =   -1  'True
            CurrentDate     =   39716.0416666667
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "No minuto da Hora"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   300
            TabIndex        =   36
            Top             =   270
            Width           =   1950
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Func. de Férias(RM->Catr.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2565
         Left            =   2827
         TabIndex        =   15
         ToolTipText     =   "Verificará os funcionarios que estaram entrando ou saindo de férias. Verifica na hora exata que você atualizou."
         Top             =   390
         Width           =   2595
         Begin VB.Frame Frame11 
            Caption         =   "Abonar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1575
            Left            =   90
            TabIndex        =   45
            Top             =   720
            Width           =   2415
            Begin VB.CheckBox CHK_Linc_Renum 
               Caption         =   "Considerar Abono p/Lic.Renumerada"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   435
               Left            =   60
               TabIndex        =   50
               Top             =   1020
               Width           =   1905
            End
            Begin VB.TextBox txt_Dia_abono 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   930
               MaxLength       =   2
               TabIndex        =   46
               Text            =   "01"
               Top             =   630
               Width           =   345
            End
            Begin MSComCtl2.DTPicker Dt_Abono 
               Height          =   345
               Left            =   930
               TabIndex        =   47
               Top             =   210
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   609
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   139460609
               CurrentDate     =   36526
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Periodo:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   30
               TabIndex        =   49
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Dias:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   60
               TabIndex        =   48
               Top             =   630
               Width           =   555
            End
         End
         Begin MSComCtl2.DTPicker txt_time_Func_ferias 
            Height          =   315
            Left            =   1050
            TabIndex        =   3
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632319
            CalendarForeColor=   16711680
            CalendarTitleBackColor=   65535
            CalendarTitleForeColor=   16711680
            CalendarTrailingForeColor=   65535
            Format          =   139460610
            UpDown          =   -1  'True
            CurrentDate     =   39716.0416666667
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Na Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   60
            TabIndex        =   16
            Top             =   390
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fun. Novatos(Rm->Catraca)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1245
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "Ajuste os horários a que horas o sistema vai fazer as novas inclusões dos novos funcionários"
         Top             =   390
         Width           =   2595
         Begin MSComCtl2.DTPicker txt_time_Func_Novo1 
            Height          =   315
            Left            =   1140
            TabIndex        =   0
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   32768
            CalendarTitleForeColor=   32768
            Format          =   139460610
            UpDown          =   -1  'True
            CurrentDate     =   39716.5208333333
         End
         Begin MSComCtl2.DTPicker txt_time_Func_Novo2 
            Height          =   315
            Left            =   1140
            TabIndex        =   1
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   32768
            Format          =   139460610
            UpDown          =   -1  'True
            CurrentDate     =   39716.9791666667
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "1a Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Left            =   90
            TabIndex        =   14
            Top             =   390
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "2a Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Left            =   90
            TabIndex        =   13
            Top             =   810
            Width           =   960
         End
      End
      Begin VB.Label LBL_MSG 
         Caption         =   "PARA ACESSO AOS PARAMETROS TECLE <ALT> + ""S"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6690
         TabIndex        =   39
         Top             =   5310
         Width           =   5925
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   13590
      Top             =   5910
   End
   Begin VB.PictureBox pichook 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13200
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   7350
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnu_taskbar 
      Caption         =   "mnu_taskbar"
      Visible         =   0   'False
      Begin VB.Menu mnu_sobre 
         Caption         =   "&Atualizar Parametros..."
      End
      Begin VB.Menu mnutraco 
         Caption         =   "-"
      End
      Begin VB.Menu mnusair 
         Caption         =   "Sa&ir"
      End
   End
End
Attribute VB_Name = "frmNotfIc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare Sub CrashVB Lib "msvbvm60" (Optional DontPassMe As Any)

Option Explicit




Public sBancoRodbel As String
Public sBancoGed As String
Public sBancoRM As String
Public sBancoVt As String
Public sBancoNFE As String
Public rs As ADODB.Recordset
Public nMinuto As Integer

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA

Private Sub CMD_ACAO_Click()

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Novatos" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Novatos
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Ferias" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Ferias
   Call Atualizacoes_Ferias_Historico
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Afastamentos" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Afastamentos
'   Call Atualizacoes_Vales_Transporte
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Desligados" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Desligados
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Mudança Setor" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Secoes
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Atestado" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Atestados
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Horas Extras Realizadas" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
'   Call Atualizacoes_Ged
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "ValeTransporte" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
'   Call Atualizacoes_Vales_Transporte
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Atualiza Funções/Setor Ged" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
'   Call Atualizacoes_Cad_Funcao_Setor_Ged
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Atualiza Horas GedXBatidas divergentes" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
'   Call Atualizacoes_Horas_Bat_Ajuste_Ged
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Email Solicitação Pendentes p/Niveis 1/2" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
'   Call Envio_Email_Nivel1_2
'   Shell "java -jar " & App.Path & "\musashi_email.jar -run", vbMinimized
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Email Solicitação Pendentes p/Niv.1/2 RH" Then
'   Call Envio_Email_Nivel1_2_RH
'   Shell "java -jar " & App.Path & "\musashi_email.jar -run", vbMinimized
End If
If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Ajuste do Status da NFE" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
'   Call Atualizacoes_Status_NFE
End If

Me.Pr_Prog.Visible = False

End Sub

Private Sub cmd_Impressao_Click()

Dim oTela As frmCristalReport
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim rs As New ADODB.Recordset
Dim sTipo As String
Dim Y As Double
Dim nx As Double

On Error GoTo Erro

Set oTela = New frmCristalReport

rs.Fields.Append "sTatus", adVarChar, 2
rs.Fields.Append "Data", adVarChar, 10
rs.Fields.Append "Hora", adVarChar, 5
rs.Fields.Append "Tipo", adVarChar, 29
rs.Fields.Append "CodFun", adVarChar, 5
rs.Fields.Append "Msg", adVarChar, 60

rs.Open

Me.MousePointer = vbHourglass

Open App.Path & "\LogRmRodBel.TXT" For Random Access Read Write As #11 Len = Len(sTexto)

Y = LOF(11) / Len(sTexto)

For nx = 1 To Y


   Get 11, nx, sTexto
   If (CDate(Mid$(sTexto.Texto, 3, 10)) >= CDate(Me.DT_Filtro_ini.Value) And _
      CDate(Mid$(sTexto.Texto, 3, 10)) <= CDate(Me.DT_Filtro_fim.Value)) And _
      (Mid$(sTexto.Texto, 14, 5) >= Mid$(DT_Hora_ini.Value, 12, 5) And _
      Mid$(sTexto.Texto, 14, 5) <= Mid$(DT_Hora_fim.Value, 12, 5)) Then
      If Me.CBO_ACOES2.ListIndex = Me.CBO_ACOES2.ListCount - 1 Or _
         Me.CBO_ACOES2.List(Me.CBO_ACOES2.ListIndex) = Trim(Mid$(sTexto.Texto, 26, 29)) Then
         rs.AddNew
         If Mid$(sTexto.Texto, 1, 1) = "0" Then
            rs.Fields("sTatus").Value = "Ok"
         ElseIf Mid$(sTexto.Texto, 1, 1) = "1" Then
            rs.Fields("sTatus").Value = "Er"
         Else
            rs.Fields("sTatus").Value = "Es"
         End If
         rs.Fields("Data").Value = Mid$(sTexto.Texto, 3, 10)
         rs.Fields("Hora").Value = Mid$(sTexto.Texto, 14, 5)
         rs.Fields("CodFun").Value = Mid$(sTexto.Texto, 20, 5)
         rs.Fields("Tipo").Value = Mid$(sTexto.Texto, 26, 29)
         rs.Fields("Msg").Value = Mid$(sTexto.Texto, 56, 60)
         rs.Update
      End If
   End If
Next

Close #11

Me.MousePointer = vbDefault

If rs.RecordCount = 0 Then
   MsgBox "Não há registros com este filtro. Altere o filtro para nova consulta."
   Exit Sub
End If

Set CrystalReport1 = Application.OpenReport(App.Path & "\RelLogRmRodbel.rpt")

rs.MoveFirst

CrystalReport1.Database.SetDataSource rs

oTela.CRViewer91.ReportSource = CrystalReport1

oTela.CRViewer91.ViewReport
oTela.CRViewer91.Refresh
rs.Clone
oTela.Show 0
oTela.CRViewer91.Refresh
Me.frm_impressao.Visible = False
Me.frm_principal.Visible = True
Me.frm_impressao.Top = 1170

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault
Close #11
Me.frm_impressao.Visible = False
Me.frm_principal.Visible = True
Me.frm_impressao.Top = 1170

End Sub

Private Sub cmd_Log_Click()
Me.FrmAcesso.Top = 6390
Me.FrmAcesso.Left = 240
If Me.frm_impressao.Visible = True Then
   Me.frm_impressao.Top = 1170
   Me.frm_impressao.Visible = False
   Me.frm_principal.Visible = True
Else
   Me.frm_impressao.Top = 1170
   Me.frm_impressao.Visible = True
   Me.frm_principal.Visible = False
End If
End Sub



Private Sub Command1_Click()
   
   Call Atualizacoes_Novatos

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 83 And Shift = 4 Then
   Me.FrmAcesso.Top = 1770
   Me.FrmAcesso.Left = 3640
   Me.FrmAcesso.Enabled = True
   Me.txt_senha.Text = ""
   Me.txt_senha.SetFocus
End If

End Sub

Private Sub Form_Load()
Dim dDate As Date
Dim cFields As Collection

Static vShowMsg As Variant 'mostra mensagem 1a vez


On Local Error Resume Next:

'If Not Empty Is Nothing Then
'On Local Error Resume Next: If Not Empty Is Nothing Then Do While Null: ReDim i(True To False) As Currency: Loop Else Debug.Assert CCur(CLng(CInt(CBool(False Imp True Xor False Eqv True)))): Stop: On Local Error GoTo 0


If App.PrevInstance Then
   MsgBox "Este Programa JÁ esta sendo processado neste computador", 16, "<ENTER>=Para Finalizar"
   Close: End
End If


If IsDate(Mid(Now(), 1, 10)) = False Then
   MsgBox "O seu computador está com o formato da DATA DIFERENTE DO PADRÃO dd/mm/yyyy. Altere as Configurações Regionais , no Painel de Controle."
   End
End If

dDate = Mid(Now(), 1, 10)
If Len(Trim(dDate)) <> 10 Then
   MsgBox "O seu computador está com o formato da DATA DIFERENTE DO PADRÃO dd/mm/yyyy. Altere as Configurações Regionais , no Painel de Controle."
   End
End If

Rem ***************************************
Rem ********  VARIAVEIS DE ACESSO A BANCO
Rem ***************************************
Call Variaveis_Acesso_Banco
Rem ***************************************

If IsEmpty(vShowMsg) Or vShowMsg = 1 Then
    MsgBox "A aplicação será reduzida a um ícone no lado direito da Barra de Tarefas do Windows.", 64, "Atualização Remota Rm -> RodBel"
    vShowMsg = 2
End If
 
t.cbSize = Len(t)
t.hWnd = pichook.hWnd
t.uId = 1&
t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
t.ucallbackMessage = WM_MOUSEMOVE
t.hIcon = Me.Icon
t.szTip = "RM_RODBEL, Comunicação Remota..." & Chr$(0) 'Texto a ser exibido quando o mouse é movido sobre o ícone.
Shell_NotifyIcon NIM_ADD, t
Me.Hide
App.TaskVisible = False

Rem CARREGAR OS COMBO E SUAS ACOES

Me.CBO_ACOES.AddItem "Novatos"
Me.CBO_ACOES.AddItem "Ferias"
Me.CBO_ACOES.AddItem "Afastamentos"
Me.CBO_ACOES.AddItem "Desligados"
Me.CBO_ACOES.AddItem "Mudança Setor"
Me.CBO_ACOES.AddItem "Mudança de Turno"
Me.CBO_ACOES.AddItem "Atestado"
Me.CBO_ACOES.AddItem "Horas Extras Realizadas"
Me.CBO_ACOES.AddItem "ValeTransporte"
Me.CBO_ACOES.AddItem "Atualiza Funções/Setor Ged"
Me.CBO_ACOES.AddItem "Atualiza Horas GedXBatidas divergentes"
Me.CBO_ACOES.AddItem "Email Solicitação Pendentes p/Niveis 1/2"
Me.CBO_ACOES.AddItem "Email Solicitação Pendentes p/Niv.1/2 RH"
Me.CBO_ACOES.AddItem "Ajuste do Status da NFE"
Me.CBO_ACOES.AddItem "TODAS AS AÇÕES ACIMA"
Me.CBO_ACOES.ListIndex = 0

Me.CBO_ACOES2.AddItem "Novatos"
Me.CBO_ACOES2.AddItem "Ferias"
Me.CBO_ACOES2.AddItem "Afastados"
Me.CBO_ACOES2.AddItem "Desligados"
Me.CBO_ACOES2.AddItem "Mudança Setor"
Me.CBO_ACOES2.AddItem "Mudança de Turno"
Me.CBO_ACOES2.AddItem "Atestado"
Me.CBO_ACOES2.AddItem "Horas Extras Realizadas"
Me.CBO_ACOES2.AddItem "ValeTransporte"
Me.CBO_ACOES2.AddItem "CadFunSet"
Me.CBO_ACOES2.AddItem "SolBatAjuste"
Me.CBO_ACOES2.AddItem "Enviar_Email"
Me.CBO_ACOES2.AddItem "Ajuste_NFE"
Me.CBO_ACOES2.AddItem "TODAS AS AÇÕES ACIMA"
Me.CBO_ACOES2.ListIndex = 0

Rem   REGISTRAR A HORA EM QUE O SISTEMA FOI INICIADO
sStatusMsg = "0"
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Sistema Ligado"
sCodFun = "0000"
sMsg = "Sistema Ligado no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & _
                   sData & ";" & _
                   sHora & ";" & _
                   sCodFun & ";" & _
                   sTipo & ";" & _
                   sMsg

Call CCTempneRegBanco.Gerar_Situacao_Log(cFields)

Me.DT_Filtro_ini.Value = "01/" & Format(Now(), "mm/yyyy")

If IsDate("28/" & Format(Now(), "mm/yyyy")) Then Me.DT_Filtro_fim.Value = "28/" & Format(Now(), "mm/yyyy")
If IsDate("29/" & Format(Now(), "mm/yyyy")) Then Me.DT_Filtro_fim.Value = "29/" & Format(Now(), "mm/yyyy")
If IsDate("30/" & Format(Now(), "mm/yyyy")) Then Me.DT_Filtro_fim.Value = "30/" & Format(Now(), "mm/yyyy")
If IsDate("31/" & Format(Now(), "mm/yyyy")) Then Me.DT_Filtro_fim.Value = "31/" & Format(Now(), "mm/yyyy")

Me.DT_Pesquisa.Value = CDate(Format(Now(), "dd/mm/yyyy")) - Val(Me.txt_Dias_Antecedencia.Text)

Set cFields = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim cFields As Collection

Rem   REGISTRAR A HORA EM QUA O SISTEMA FOI DESLIGADO
sStatusMsg = "0"
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Sistema Desligado"
sCodFun = "0000"
sMsg = "Sistema Desligado no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & _
                   sData & ";" & _
                   sHora & ";" & _
                   sCodFun & ";" & _
                   sTipo & ";" & _
                   sMsg

Call CCTempneRegBanco.Gerar_Situacao_Log(cFields)
    
t.cbSize = Len(t)
t.hWnd = pichook.hWnd
t.uId = 1&
Shell_NotifyIcon NIM_DELETE, t  'Remove o ícone da barra de tarefas.
    
End Sub

Private Sub Form_Resize()
    If (Me.WindowState) = 1 Then
        Me.Hide
    End If
End Sub

Private Sub mnu_sobre_Click()
    Me.WindowState = 0
    Me.Show
End Sub

Private Sub mnusair_Click()
    Unload Me
    End
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'pichook é uma picture box, utilizada pelo Windows para
'reconhecer o ícone na barra de tarefas.
    Static rec As Boolean, msg As Long
    msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case msg
            Case WM_LBUTTONDBLCLK:
                 Me.PopupMenu mnu_taskbar
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
        'Se for pressionado o botão direito
        'sobre o ícone, é exibido um menu pop-up.
                Me.PopupMenu mnu_taskbar    'mnuBar-menu criado no form.
        End Select
        rec = False
    End If
End Sub

Private Sub Timer1_Timer()

Dim tTempo As String
Dim tTempo2 As String

tTempo = Format(Now(), "hh:mm:ss")
tTempo2 = Format(Now(), "hh:mm:ss")

Rem AQUI MARCOS
Rem Exit Sub


Rem ***************************************
Rem ATUALIZAR A DATA DO PERIODO DE PESQUISA
Rem ***************************************
    If Val(Me.txt_Dias_Antecedencia.Text) > 0 Then
       Me.DT_Pesquisa.Value = CDate(Format(Now(), "dd/mm/yyyy")) - Val(Me.txt_Dias_Antecedencia.Text)
    Else
       Me.DT_Pesquisa.Value = CDate(Format(Now(), "dd/mm/yyyy"))
    End If
Rem ***************************************

Rem ****************************************************************
Rem verificar tempos para Inclusao de novos funcionários
Rem ****************************************************************
    If tTempo = Format(txt_time_Func_Novo1.Value, "hh:mm:ss") Then
        Call Atualizacoes_Novatos
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If

    If tTempo = Format(txt_time_Func_Novo2.Value, "hh:mm:ss") Then
        Call Atualizacoes_Novatos
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem Verificar tempos para demissao dos funcionários
Rem ****************************************************************
    If Mid$(tTempo, 4, 5) = Mid$(Format(txt_time_Func_Desligados.Value, "hh:mm:ss"), 4, 5) Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Desligados.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Desligados.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Desligados
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem verificar tempos para FERIAS dos funcionários
Rem ****************************************************************
    If tTempo = Format(txt_time_Func_ferias.Value, "hh:mm:ss") Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_ferias.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_ferias.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Ferias
        Call Atualizacoes_Ferias_Historico
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem verificar tempos para Mudanca de Secao dos funcionários
Rem ****************************************************************
    If Mid$(tTempo, 4, 5) = Mid$(Format(txt_time_Func_Muda_Setor.Value, "hh:mm:ss"), 4, 5) Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Muda_Setor.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Muda_Setor.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Secoes
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem verificar tempos para Afastamentos de funcionários
Rem ****************************************************************
    If tTempo = Format(txt_time_Func_Afast_Ini.Value, "hh:mm:ss") Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Afast_Ini.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Afast_Ini.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Afastamentos
'        Call Atualizacoes_Vales_Transporte
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If

    If tTempo = Format(txt_time_Func_Afast_Fim.Value, "hh:mm:ss") Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Afast_Fim.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Afast_Fim.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Afastamentos
'        Call Atualizacoes_Vales_Transporte
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem verificar tempos para Atestados de funcionários
Rem ****************************************************************
    If tTempo = Format(txt_time_Func_Atest_Ini.Value, "hh:mm:ss") Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Atest_Ini.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Atest_Ini.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Atestados
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If

    If tTempo = Format(txt_time_Func_Atest_Fim.Value, "hh:mm:ss") Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Atest_Fim.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Atest_Fim.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Atestados
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If

Rem ****************************************************************
Rem verificar tempos para Mudanca do statusda NFE
Rem ****************************************************************
    If Mid$(tTempo, 4, 5) = Mid$(Format(txt_time_Hora_Status_NFE.Value, "hh:mm:ss"), 4, 5) Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Hora_Status_NFE.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Hora_Status_NFE.Value, "hh:mm:ss"), 4, 5)) Then
'        Call Atualizacoes_Status_NFE
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem verificar tempos para Atualização das Horas extras Batidas e não solicitadas para atualizar o GED.
Rem ****************************************************************
    If tTempo = Format(txt_time_Hora_Real_Ged.Value, "hh:mm:ss") Then
'    If Mid$(tTempo, 4, 5) = Mid$(Format(txt_time_Hora_Real_Ged.Value, "hh:mm:ss"), 4, 5) Or _
'       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Hora_Real_Ged.Value, "hh:mm:ss"), 4, 5) And _
'        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Hora_Real_Ged.Value, "hh:mm:ss"), 4, 5)) Then
'        Call Atualizacoes_Ged
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************


Rem ****************************************************************
Rem verificar tempos para Mudanca de Secao dos funcionários
Rem ****************************************************************
    If Mid$(tTempo, 4, 5) = Mid$(Format(txt_time_Cad_Funcao_Secao_Ged.Value, "hh:mm:ss"), 4, 5) Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Cad_Funcao_Secao_Ged.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Cad_Funcao_Secao_Ged.Value, "hh:mm:ss"), 4, 5)) Then
'        Call Atualizacoes_Cad_Funcao_Setor_Ged
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If

Rem ****************************************************************
Rem verificar tempos para Atualização das Batidas ajustadas para atualizar o GED já atualizado.
Rem ****************************************************************
    If tTempo = Format(txt_time_Hora_Bat_Ponto.Value, "hh:mm:ss") Then
'    If Mid$(tTempo, 4, 5) = Mid$(Format(txt_time_Hora_Real_Ged.Value, "hh:mm:ss"), 4, 5) Or _
'       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Hora_Real_Ged.Value, "hh:mm:ss"), 4, 5) And _
'        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Hora_Real_Ged.Value, "hh:mm:ss"), 4, 5)) Then
'        Call Atualizacoes_Horas_Bat_Ajuste_Ged
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************


Rem ****************************************************************
Rem verificar existência de Solicitações pendentes ou não solicitadas e envia emails para os niveis 1 e 2
Rem ****************************************************************
    If tTempo = Format(txt_time_Func_HE_Ext_Nautor.Value, "hh:mm:ss") Then
'        Call Envio_Email_Nivel1_2
'        Shell "java -jar " & App.Path & "\musashi_email.jar -run", vbMinimized
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If

Rem ****************************************************************
Rem verificar existência de Solicitações pendentes ou não solicitadas
Rem e envia emails para os niveis 1 e 2 do setor 651(RH), NO DIA E HORA MARCADA
Rem ****************************************************************
    If tTempo = Format(txt_time_Func_HE_Ext_Nautor_RH.Value, "hh:mm:ss") And _
        Format(Now(), "DD") = Trim(Me.txt_dia_email.Text) Then
'        Call Envio_Email_Nivel1_2_RH
'        Shell "java -jar " & App.Path & "\musashi_email.jar -run", vbMinimized
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If

Me.Pr_Prog.Visible = False

Rem ****************************************************
Rem CONTROLE DE MENSAGEM PISCANTE E MENSAGEM A SER DITA
Rem ****************************************************

If Me.LBL_MSG.Visible = True Then
   Me.LBL_MSG.Visible = False
Else
   Me.LBL_MSG.Visible = True
End If
Rem ****************************************************

Rem ****************************************************
Rem CONTROLE DE MENSAGEM PISCANTE E MENSAGEM A SER DITA
Rem ****************************************************
If Me.frm_principal.Enabled = True Then
   nMinuto = nMinuto - 1
   If nMinuto < 0 Then
      Me.frm_principal.Enabled = False
   End If
   Me.LBL_MSG.Caption = "ACESSO LIBERADO ATUALIZE OS PARAMETROS.(T) " & Str(nMinuto) & " s."
'   Me.LBL_MSG.ForeColor = &H8000&
Else
   Me.LBL_MSG.Caption = "PARA ACESSO AOS PARAMETROS TECLE <ALT> + 'S'"
'   Me.LBL_MSG.ForeColor = &H8000000F
End If
Rem ****************************************************

'Me.Caption = "Atualização Remota Rm -> RodBel - Data Sistema " & Format(Now(), "dd/mm/yyyy") & " Hora : " & Format(Now(), "hh:mm:ss")
Me.Caption = "Atualização Remota Rm -> RodBel - V.1.062012, Data Sistema " & Format(Now(), "dd/mm/yyyy")

End Sub

Private Sub Variaveis_Acesso_Banco()
Dim Nada As String
Dim nx As Integer
Dim filenum As Integer

'On Error GoTo Erro

Rem teste da data do computador para o formato dd/mm/yyyy

If IsDate(Mid(Now(), 1, 10)) = False Then
   MsgBox "O seu computador está com o formato da DATA DIFERENTE DO PADRÃO dd/mm/yyyy. Altere as Configurações Regionais , no Painel de Controle."
   End
End If

If Len(Trim(Mid(Now(), 1, 10))) <> 10 Then
   MsgBox "O seu computador está com o formato da DATA DIFERENTE DO PADRÃO dd/mm/yyyy. Altere as Configurações Regionais , no Painel de Controle."
   End
End If
 
Nada = App.Path & "\LOCALIZA.TXT"
If Dir$(Nada) = "" Then
   MsgBox "Arquivo de inicialização não encontrado, Procure o responsável!", 16, "Programa Cancelado"
   End
End If
 
filenum = FreeFile
Open App.Path & "\localiza.txt" For Input Shared As filenum
Input #filenum, sBancoRodbel        'Caminho do banco da rodbel
Input #filenum, sBancoRM            'caminho do rm
Input #filenum, sBancoGed           'caminho do ged
Input #filenum, sBancoVt           'caminho do vale transporte
Input #filenum, sBancoNFE           'caminho do nfe
Close filenum

sBancoRodbel = Trim(retirarComentario(sBancoRodbel))
sBancoRM = Trim(retirarComentario(sBancoRM))
sBancoGed = Trim(retirarComentario(sBancoGed))
sBancoVt = Trim(retirarComentario(sBancoVt))
sBancoNFE = Trim(retirarComentario(sBancoNFE))

Rem *************************  A T E N Ç Ã O *****************************************
Rem *************************  BASE TESTE    *****************************************
Rem **********************************************************************************
' sBancoRodbel = "Provider=SQLOLEDB.1;" & _
'                 "Password=F396B50;" & _
'                 "Persist Security Info=True;" & _
'                 "User ID=sa;" & _
'                 "Initial Catalog=CopiaRBACESSO;" & _
'                 "Data Source=msb-25"
'
' Me.Caption = &H8000000A
'
' sBancoRM = "Provider=SQLOLEDB.1;" & _
'                 "Password=F396B50;" & _
'                 "Persist Security Info=True;" & _
'                 "User ID=sa;" & _
'                 "Initial Catalog=CopiaRM_Recife;" & _
'                 "Data Source=msb-25"
'
' sBancoGed = "Provider=SQLOLEDB.1;" & _
'                 "Password=F396B50;" & _
'                 "Persist Security Info=True;" & _
'                 "User ID=sa;" & _
'                 "Initial Catalog=Copia_Ged_Musashi;" & _
'                 "Data Source=msb-25"
'
' sBancoVt = "Provider=SQLOLEDB.1;" & _
'                 "Password=F396B50;" & _
'                 "Persist Security Info=True;" & _
'                 "User ID=sa;" & _
'                 "Initial Catalog=CopiaValetrp;" & _
'                 "Data Source=msb-25"
' sBancoNFE = "Provider=SQLOLEDB.1;" & _
'                 "Password=F396B50;" & _
'                 "Persist Security Info=True;" & _
'                 "User ID=sa;" & _
'                 "Initial Catalog=NFEQA;" & _
'                 "Data Source=msb-25"



Rem *************************  A T E N Ç Ã O *****************************************
Rem *************************  BASE PRODUCAO *****************************************
Rem **********************************************************************************

''' sBancoRodbel = "Provider=SQLOLEDB.1;" & _
'''                 "Password=F396B50;" & _
'''                 "Persist Security Info=True;" & _
'''                 "User ID=sa;" & _
'''                 "Initial Catalog=MDACESSO;" & _
'''                 "Data Source=msb-25"
'''
''' Me.Caption = &H8000000A
'''
''' sBancoRM = "Provider=SQLOLEDB.1;" & _
'''                 "Password=F396B50;" & _
'''                 "Persist Security Info=True;" & _
'''                 "User ID=sa;" & _
'''                 "Initial Catalog=CorporeRM;" & _
'''                 "Data Source=msb-25"
'''
''' sBancoGed = "Provider=SQLOLEDB.1;" & _
'''                 "Password=F396B50;" & _
'''                 "Persist Security Info=True;" & _
'''                 "User ID=sa;" & _
'''                 "Initial Catalog=Ged_Musashi;" & _
'''                 "Data Source=msb-25"
'''
''' sBancoVt = "Provider=SQLOLEDB.1;" & _
'''                 "Password=F396B50;" & _
'''                 "Persist Security Info=True;" & _
'''                 "User ID=sa;" & _
'''                 "Initial Catalog=Valetrp;" & _
'''                 "Data Source=msb-25"
'''
''' sBancoNFE = "Provider=SQLOLEDB.1;" & _
'''                 "Password=F396B50;" & _
'''                 "Persist Security Info=True;" & _
'''                 "User ID=sa;" & _
'''                 "Initial Catalog=NFE;" & _
'''                 "Data Source=msb-25"

Rem **********************************************************************************
Rem *************************  A T E N Ç Ã O *****************************************
Rem *************************  A T E N Ç Ã O *****************************************


End Sub

Private Sub Atualizacoes_Novatos()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(DT_Pesquisa.Value, "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Funcionarios_Novatos(sBancoRM, sBancoRodbel, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Desligados()
Dim nx As Double
Dim nLinhas As Double
Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(DT_Pesquisa.Value, "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Funcionarios_Desligados(sBancoRM, sBancoRodbel, sDataAdm)
Set rs = CCTempneRegBanco.Funcionarios_AfastLongoTempo(sBancoRM, sBancoRodbel, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Ferias_Historico()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(DT_Pesquisa.Value, "dd/mm/yyyy")) - 30

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Funcionarios_Ferias_Histo(sBancoRM, sBancoRodbel, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub

Private Sub Atualizacoes_Ferias()

Dim sDataAdm As String
Dim sDataAbn As String
Dim sAbnDia As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(DT_Pesquisa.Value, "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"
sDataAbn = IIf(Me.Dt_Abono.Value = "__/__/____", "", Me.Dt_Abono.Value)
sAbnDia = Val(Me.txt_Dia_abono.Text)

Set rs = CCTempneRegBanco.Funcionarios_Ferias(sBancoRM, sBancoRodbel, sDataAdm, sDataAbn, Val(sAbnDia))

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Secoes()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(DT_Pesquisa.Value, "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Funcionario_Historico_Secao(sBancoRM, sBancoRodbel, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Afastamentos()

Dim sDataAdm As String
Dim sDataAbn As String
Dim sAbnDia As Integer

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(Now(), "dd/mm/yyyy"))
sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

If Me.CHK_Linc_Renum.Value = 1 Then
   sDataAbn = IIf(Me.Dt_Abono.Value = "__/__/____", "", Me.Dt_Abono.Value)
   sAbnDia = Val(Me.txt_Dia_abono.Text)
Else
   sDataAbn = IIf(Me.Dt_Abono.Value = "__/__/____", "", Me.Dt_Abono.Value)
   sAbnDia = 0
End If

Set rs = CCTempneRegBanco.Funcionarios_Afastado(sBancoRM, sBancoRodbel, sDataAdm, sDataAbn, sAbnDia)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Vales_Transporte()
Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(DT_Pesquisa.Value, "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

'Set rs = CCTempneRegBanco.Funcionarios_Novatos(sBancoRM, sBancoRodbel, sDataAdm)

Set rs = CCTempneRegBanco.Atualizar_Vale_Transporte(sBancoRM, sBancoGed, sBancoVt, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault

End Sub

Private Sub Atualizacoes_Ged()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(Now(), "dd/mm/yyyy"))
'sDataAdm = CDate(Format(DT_Pesquisa.Value, "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Atualizar_Horas_Extras_Ged(sBancoRM, sBancoGed)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Cad_Funcao_Setor_Ged()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(Now(), "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Atualizar_Cad_Funcao_Setor_Ged(sBancoRM, sBancoGed, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Status_NFE()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = Format(DateAdd("h", -2, Now()), "YYYY-MM-DD HH:MM:SS")

Set rs = CCTempneRegBanco.Atualizar_Status_NFE(sBancoNFE, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Horas_Bat_Ajuste_Ged()
Rem esta função fará as atualizaçoes das solicitações geradas pela ausência de solicitaçao. No caso serão realizadas
Rem comparações com as horas já cadastradas na solicitação com o registro de ponto, caso esteja diterente, atualizar com a batida.

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(Now(), "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Atualizar_Horas_Bat_Ajuste_Ged(sBancoRM, sBancoGed, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub

Private Sub Envio_Email_Nivel1_2()
Rem esta função fará as atualizaçoes das solicitações geradas pela ausência de solicitaçao. No caso serão realizadas
Rem comparações com as horas já cadastradas na solicitação com o registro de ponto, caso esteja diterente, atualizar com a batida.

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(Now(), "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Enviar_Email_TresdiasUteis_Ged(sBancoRM, sBancoGed, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Envio_Email_Nivel1_2_RH()
Rem esta função fará as atualizaçoes das solicitações geradas pela ausência de solicitaçao. No caso serão realizadas
Rem comparações com as horas já cadastradas na solicitação com o registro de ponto, caso esteja diterente, atualizar com a batida.

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(Now(), "dd/mm/yyyy"))

'sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Format(Me.txt_dia_email.Text, "00") & "'"

Set rs = CCTempneRegBanco.Enviar_Email_RH(sBancoRM, sBancoGed, sDataAdm)

Set rs = CCTempneRegBanco.Enviar_Email_N1N2(sBancoRM, sBancoGed, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub

Private Sub Atualizacoes_Atestados()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(Now(), "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Funcionarios_Atestado(sBancoRM, sBancoRodbel, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub

Public Function CCTempneRegBanco() As neRegBanco
     Set CCTempneRegBanco = New neRegBanco
End Function

Private Sub txt_Dias_Antecedencia_Change()
If Val(Me.txt_Dias_Antecedencia.Text) > 0 Then
   Me.DT_Pesquisa.Value = CDate(Format(Now(), "dd/mm/yyyy")) - Val(Me.txt_Dias_Antecedencia.Text)
Else
   Me.DT_Pesquisa.Value = CDate(Format(Now(), "dd/mm/yyyy"))
End If

End Sub

Private Sub txt_senha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
   If UCase(Me.txt_senha.Text) <> Format(Now(), "YYYYMMDD") & "ACS" Then
      MsgBox "SENHA NÃO CONFERE, TENTE NOVAMENTE OU TECLE <ESC> "
      Me.txt_senha.Text = ""
      Me.txt_senha.SetFocus
   Else
      MsgBox "Você terá 10 (Dez) minutos para atualização dos parametros"
      Me.frm_principal.Enabled = True
      Me.txt_senha.Text = ""
      Me.frm_principal.Enabled = True
      Me.FrmAcesso.Top = 6390
      Me.FrmAcesso.Left = 240
      nMinuto = 600
   End If
End If

If KeyAscii = 27 Then
   Me.frm_principal.Enabled = False
   Me.FrmAcesso.Top = 6390
   Me.FrmAcesso.Left = 240
   Me.txt_senha.Text = ""
   Me.frm_principal.Enabled = False
End If
   

End Sub

