VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1C90141-3FBE-4464-B25B-D4CA17FB66F3}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmOption 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ϵͳ����"
   ClientHeight    =   6450
   ClientLeft      =   3975
   ClientTop       =   2985
   ClientWidth     =   9330
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9330
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "���в���"
      TabPicture(0)   =   "frmOption.frx":1CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "���Ʋ���"
      TabPicture(1)   =   "frmOption.frx":1D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "�ֹ�ά��"
      TabPicture(2)   =   "frmOption.frx":1D32
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "TPMS����������"
      TabPicture(3)   =   "frmOption.frx":1D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame13"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         Height          =   5925
         Left            =   -74940
         TabIndex        =   59
         Top             =   360
         Width           =   9105
         Begin VB.Frame Frame18 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��Ͻ����ӡ����"
            Height          =   765
            Left            =   90
            TabIndex        =   91
            Top             =   5070
            Width           =   8895
            Begin VB.CheckBox chkOnlyPrintNGWriteResult 
               BackColor       =   &H00FFFFFF&
               Caption         =   "chkPrintNGResult"
               Height          =   345
               Left            =   2010
               TabIndex        =   95
               Top             =   270
               Width           =   195
            End
            Begin VB.CheckBox chkPrintNGFlow 
               BackColor       =   &H00FFFFFF&
               Caption         =   "checkNGFlow"
               Height          =   345
               Left            =   4320
               TabIndex        =   94
               Top             =   270
               Width           =   195
            End
            Begin VB.CommandButton Command7 
               Caption         =   "�ֶ���ӡ"
               Height          =   375
               Left            =   7380
               TabIndex        =   93
               Top             =   270
               Width           =   1305
            End
            Begin VB.TextBox txtVIN 
               Height          =   315
               Left            =   5280
               TabIndex        =   92
               Top             =   300
               Width           =   2025
            End
            Begin VB.Label Label31 
               BackColor       =   &H00FFFFFF&
               Caption         =   "����ӡNG����Ͻ����"
               Height          =   225
               Left            =   180
               TabIndex        =   98
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label32 
               BackColor       =   &H00FFFFFF&
               Caption         =   "VIN��"
               Height          =   225
               Left            =   4800
               TabIndex        =   97
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label30 
               BackColor       =   &H00FFFFFF&
               Caption         =   "����ӡNG��������̣�"
               Height          =   225
               Left            =   2520
               TabIndex        =   96
               Top             =   360
               Width           =   1815
            End
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ʼλ����      "
            Height          =   1635
            Left            =   90
            TabIndex        =   61
            Top             =   3360
            Width           =   8895
            Begin VB.CommandButton btMTOCModi 
               Caption         =   "�޸�"
               Height          =   375
               Left            =   210
               TabIndex        =   77
               Top             =   1110
               Width           =   1515
            End
            Begin VB.TextBox txtMTOCLen 
               Height          =   315
               Left            =   1170
               TabIndex        =   74
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox txtMtocStartIndex 
               Height          =   315
               Left            =   1170
               TabIndex        =   72
               Top             =   300
               Width           =   1695
            End
            Begin VB.Label Label24 
               BackColor       =   &H00FFFFFF&
               Caption         =   "��ע��MTOC��ΪBF1B-FM6-00-B1-V������λ��Ϊ3����������ΪFM6��"
               Height          =   225
               Left            =   3060
               TabIndex        =   76
               Top             =   750
               Width           =   5655
            End
            Begin VB.Label Label23 
               BackColor       =   &H00FFFFFF&
               Caption         =   "��ע��MTOC��ΪBF1B-FM6-00-B1-V����ʼλΪ6����ӵڶ���F��ʼ��"
               Height          =   225
               Left            =   3060
               TabIndex        =   75
               Top             =   360
               Width           =   5655
            End
            Begin VB.Label Label22 
               BackColor       =   &H00FFFFFF&
               Caption         =   "�����볤��"
               Height          =   225
               Left            =   240
               TabIndex        =   73
               Top             =   750
               Width           =   915
            End
            Begin VB.Label Label21 
               BackColor       =   &H00FFFFFF&
               Caption         =   "��ʼλ�ã�"
               Height          =   225
               Left            =   240
               TabIndex        =   71
               Top             =   360
               Width           =   915
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�������б�      "
            Height          =   3105
            Left            =   60
            TabIndex        =   60
            Top             =   180
            Width           =   8925
            Begin VB.Frame Frame16 
               BackColor       =   &H00FFFFFF&
               Height          =   2595
               Left            =   4830
               TabIndex        =   63
               Top             =   240
               Width           =   3915
               Begin VB.CommandButton btTPMSCancle 
                  Caption         =   "ȡ��"
                  Height          =   375
                  Left            =   2220
                  TabIndex        =   78
                  Top             =   1830
                  Width           =   1095
               End
               Begin VB.CommandButton btTPMSDel 
                  Caption         =   "ɾ��"
                  Height          =   375
                  Left            =   420
                  TabIndex        =   70
                  Top             =   1830
                  Width           =   1095
               End
               Begin VB.CommandButton btTPMSModi 
                  Caption         =   "�޸�"
                  Height          =   375
                  Left            =   2190
                  TabIndex        =   69
                  Top             =   1230
                  Width           =   1095
               End
               Begin VB.TextBox txtTPMSCode 
                  Height          =   315
                  Left            =   1230
                  TabIndex        =   67
                  Top             =   720
                  Width           =   2235
               End
               Begin VB.TextBox txtTPMSID 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1230
                  Locked          =   -1  'True
                  TabIndex        =   65
                  Top             =   270
                  Width           =   2235
               End
               Begin VB.CommandButton btTPMSAdd 
                  Caption         =   "����"
                  Height          =   375
                  Left            =   420
                  TabIndex        =   64
                  Top             =   1230
                  Width           =   1095
               End
               Begin VB.Label Label20 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "�����룺"
                  Height          =   225
                  Left            =   300
                  TabIndex        =   68
                  Top             =   750
                  Width           =   735
               End
               Begin VB.Label Label19 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "��    �ţ�"
                  Height          =   225
                  Left            =   300
                  TabIndex        =   66
                  Top             =   300
                  Width           =   735
               End
            End
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
               Height          =   2505
               Left            =   210
               TabIndex        =   62
               Top             =   360
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   4419
               _Version        =   393216
               BackColor       =   16777215
               BackColorFixed  =   -2147483639
               BackColorBkg    =   16777215
               Appearance      =   0
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Height          =   5925
         Left            =   90
         TabIndex        =   33
         Top             =   360
         Width           =   9075
         Begin VB.Frame Frame17 
            BackColor       =   &H00FFFFFF&
            Caption         =   "������ģʽ�趨         "
            Height          =   795
            Left            =   150
            TabIndex        =   86
            Top             =   2310
            Width           =   8625
            Begin VB.CommandButton cmdMdlSave 
               Caption         =   "����"
               Height          =   345
               Left            =   5790
               TabIndex        =   88
               Top             =   300
               Width           =   1425
            End
            Begin VB.TextBox txtMdl 
               Height          =   315
               Left            =   1140
               TabIndex        =   87
               Top             =   330
               Width           =   3585
            End
            Begin VB.Label Label29 
               BackColor       =   &H00FFFFFF&
               Caption         =   "ģ  ʽ��"
               Height          =   225
               Left            =   390
               TabIndex        =   90
               Top             =   390
               Width           =   885
            End
            Begin VB.Label Label28 
               BackColor       =   &H00FFFFFF&
               Caption         =   "(���ŷָ�)"
               Height          =   225
               Left            =   4740
               TabIndex        =   89
               Top             =   390
               Width           =   915
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00FFFFFF&
            Caption         =   "������ѹ��ֵ��Χ�趨            "
            Height          =   765
            Left            =   150
            TabIndex        =   39
            Top             =   3240
            Width           =   8625
            Begin VB.CommandButton cmdPreSave 
               Caption         =   "����"
               Height          =   345
               Left            =   5790
               TabIndex        =   44
               Top             =   270
               Width           =   1425
            End
            Begin VB.TextBox txtPreMax 
               Height          =   315
               Left            =   4080
               TabIndex        =   43
               Top             =   300
               Width           =   1515
            End
            Begin VB.TextBox txtPreMin 
               Height          =   315
               Left            =   1140
               TabIndex        =   41
               Top             =   300
               Width           =   1515
            End
            Begin VB.Label Label16 
               BackColor       =   &H00FFFFFF&
               Caption         =   "���ֵ��"
               Height          =   225
               Left            =   3330
               TabIndex        =   42
               Top             =   360
               Width           =   885
            End
            Begin VB.Label Label13 
               BackColor       =   &H00FFFFFF&
               Caption         =   "��Сֵ��"
               Height          =   225
               Left            =   390
               TabIndex        =   40
               Top             =   360
               Width           =   885
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "���������ٶ�ֵ��Χ�趨            "
            Height          =   765
            Left            =   180
            TabIndex        =   51
            Top             =   5040
            Width           =   8595
            Begin VB.CommandButton cmdAcSpeedSave 
               Caption         =   "����"
               Height          =   345
               Left            =   5790
               TabIndex        =   54
               Top             =   240
               Width           =   1425
            End
            Begin VB.TextBox txtAcSpeedMax 
               Height          =   315
               Left            =   4080
               TabIndex        =   53
               Top             =   270
               Width           =   1515
            End
            Begin VB.TextBox txtAcSpeedMin 
               Height          =   315
               Left            =   1110
               TabIndex        =   52
               Top             =   270
               Width           =   1515
            End
            Begin VB.Label Label18 
               BackColor       =   &H00FFFFFF&
               Caption         =   "���ֵ��"
               Height          =   225
               Left            =   3300
               TabIndex        =   56
               Top             =   330
               Width           =   885
            End
            Begin VB.Label Label17 
               BackColor       =   &H00FFFFFF&
               Caption         =   "��Сֵ��"
               Height          =   225
               Left            =   390
               TabIndex        =   55
               Top             =   330
               Width           =   885
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�������¶�ֵ��Χ�趨            "
            Height          =   765
            Left            =   180
            TabIndex        =   45
            Top             =   4140
            Width           =   8595
            Begin VB.TextBox txtTempMin 
               Height          =   315
               Left            =   1110
               TabIndex        =   48
               Top             =   300
               Width           =   1515
            End
            Begin VB.TextBox txtTempMax 
               Height          =   315
               Left            =   4080
               TabIndex        =   47
               Top             =   300
               Width           =   1515
            End
            Begin VB.CommandButton cmdTempSave 
               Caption         =   "����"
               Height          =   345
               Left            =   5760
               TabIndex        =   46
               Top             =   270
               Width           =   1425
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFFF&
               Caption         =   "��Сֵ��"
               Height          =   225
               Left            =   390
               TabIndex        =   50
               Top             =   360
               Width           =   885
            End
            Begin VB.Label Label14 
               BackColor       =   &H00FFFFFF&
               Caption         =   "���ֵ��"
               Height          =   225
               Left            =   3300
               TabIndex        =   49
               Top             =   360
               Width           =   885
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�Ų����м���ģʽ    "
            Height          =   765
            Left            =   210
            TabIndex        =   38
            Top             =   180
            Width           =   8565
            Begin VB.CheckBox chkOnlyScanVINCode 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Check1"
               Height          =   345
               Left            =   5430
               TabIndex        =   84
               Top             =   180
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.CheckBox chkAllQueue 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Check1"
               Height          =   345
               Left            =   330
               TabIndex        =   82
               Top             =   300
               Width           =   195
            End
            Begin VB.CommandButton Command10 
               Caption         =   "�鿴�Ų���������"
               Height          =   400
               Left            =   5010
               TabIndex        =   80
               Top             =   150
               Visible         =   0   'False
               Width           =   285
            End
            Begin VB.CommandButton Command5 
               Caption         =   "�ֶ������Ų���������"
               Height          =   400
               Left            =   4500
               TabIndex        =   79
               Top             =   150
               Visible         =   0   'False
               Width           =   405
            End
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "��ɨ��VIN�룬��MESȡMTOC��"
               Height          =   225
               Left            =   5700
               TabIndex        =   85
               Top             =   240
               Visible         =   0   'False
               Width           =   2565
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "У���Ų�������Ϣ"
               Height          =   270
               Left            =   600
               TabIndex        =   81
               Top             =   360
               Width           =   1800
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�޸Ĺ�������      "
            Height          =   1185
            Left            =   180
            TabIndex        =   34
            Top             =   1020
            Width           =   8595
            Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
               Left            =   7320
               Top             =   360
               _ExtentX        =   6588
               _ExtentY        =   1085
               ColorScheme     =   2
               Common_Dialog   =   0   'False
            End
            Begin VB.TextBox Text2 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   2280
               PasswordChar    =   "*"
               TabIndex        =   58
               Top             =   660
               Width           =   2055
            End
            Begin VB.TextBox Text1 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   2280
               PasswordChar    =   "*"
               TabIndex        =   57
               Top             =   240
               Width           =   2055
            End
            Begin VB.CommandButton Command6 
               Caption         =   "����������"
               Height          =   375
               Left            =   4500
               TabIndex        =   35
               Top             =   630
               Width           =   1455
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "�����������룺"
               Height          =   270
               Left            =   600
               TabIndex        =   37
               Top             =   300
               Width           =   1800
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "���ٴ����������룺"
               Height          =   270
               Left            =   600
               TabIndex        =   36
               Top             =   735
               Width           =   1800
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   5955
         Left            =   -74940
         TabIndex        =   17
         Top             =   360
         Width           =   9135
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�޸�  "
            Height          =   5595
            Left            =   5700
            TabIndex        =   22
            Top             =   270
            Width           =   3345
            Begin VB.TextBox txtGroupCtrl 
               Height          =   315
               Left            =   1050
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   690
               Width           =   2085
            End
            Begin VB.TextBox txtKeyCtrl 
               Height          =   315
               Left            =   1050
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   1410
               Width           =   2085
            End
            Begin VB.TextBox txtDescriptionCtrl 
               Height          =   765
               Left            =   1050
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   2070
               Width           =   2085
            End
            Begin VB.TextBox txtValueCtrl 
               Height          =   315
               Left            =   1050
               TabIndex        =   25
               Top             =   3210
               Width           =   2085
            End
            Begin VB.CommandButton Command4 
               Caption         =   "ȡ��"
               Height          =   375
               Left            =   450
               TabIndex        =   24
               Top             =   4410
               Width           =   1095
            End
            Begin VB.CommandButton Command3 
               Caption         =   "�޸�"
               Height          =   375
               Left            =   2040
               TabIndex        =   23
               Top             =   4410
               Width           =   1095
            End
            Begin VB.Label Label10 
               BackColor       =   &H00FFFFFF&
               Caption         =   "�飺"
               Height          =   225
               Left            =   660
               TabIndex        =   32
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label9 
               BackColor       =   &H00FFFFFF&
               Caption         =   "�ؼ��֣�"
               Height          =   255
               Left            =   300
               TabIndex        =   31
               Top             =   1440
               Width           =   765
            End
            Begin VB.Label Label8 
               BackColor       =   &H00FFFFFF&
               Caption         =   "������"
               Height          =   255
               Left            =   480
               TabIndex        =   30
               Top             =   2280
               Width           =   585
            End
            Begin VB.Label Label7 
               BackColor       =   &H00FFFFFF&
               Caption         =   "ֵ��"
               Height          =   255
               Left            =   660
               TabIndex        =   29
               Top             =   3240
               Width           =   735
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�����б�    "
            Height          =   5595
            Left            =   90
            TabIndex        =   18
            Top             =   270
            Width           =   5505
            Begin VB.ComboBox ComboCtrl 
               Height          =   315
               Left            =   1650
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   270
               Width           =   2505
            End
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
               Height          =   4725
               Left            =   60
               TabIndex        =   20
               Top             =   750
               Width           =   5355
               _ExtentX        =   9446
               _ExtentY        =   8334
               _Version        =   393216
               Cols            =   4
               BackColor       =   16777215
               BackColorFixed  =   -2147483639
               BackColorBkg    =   16777215
               Appearance      =   0
            End
            Begin VB.Label Label6 
               BackColor       =   &H00FFFFFF&
               Caption         =   "�飺"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1260
               TabIndex        =   21
               Top             =   330
               Width           =   825
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   5955
         Left            =   -74940
         TabIndex        =   1
         Top             =   360
         Width           =   9135
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�����б�    "
            Height          =   5595
            Left            =   90
            TabIndex        =   13
            Top             =   270
            Width           =   5505
            Begin VB.ComboBox ComboRun 
               Height          =   315
               Left            =   1650
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   270
               Width           =   2565
            End
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
               Height          =   4725
               Left            =   60
               TabIndex        =   14
               Top             =   750
               Width           =   5325
               _ExtentX        =   9393
               _ExtentY        =   8334
               _Version        =   393216
               Cols            =   4
               BackColor       =   16777215
               BackColorFixed  =   -2147483639
               BackColorBkg    =   16777215
               Appearance      =   0
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFFFFF&
               Caption         =   "�飺"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1230
               TabIndex        =   16
               Top             =   330
               Width           =   825
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�޸�  "
            Height          =   5595
            Left            =   5700
            TabIndex        =   2
            Top             =   270
            Width           =   3345
            Begin VB.CommandButton Command2 
               Caption         =   "�޸�"
               Height          =   375
               Left            =   2040
               TabIndex        =   12
               Top             =   4350
               Width           =   1095
            End
            Begin VB.CommandButton Command1 
               Caption         =   "ȡ��"
               Height          =   375
               Left            =   450
               TabIndex        =   11
               Top             =   4350
               Width           =   1095
            End
            Begin VB.TextBox txtValueRun 
               Height          =   315
               Left            =   1020
               TabIndex        =   10
               Top             =   3300
               Width           =   2085
            End
            Begin VB.TextBox txtDescriptionRun 
               Height          =   765
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   2070
               Width           =   2085
            End
            Begin VB.TextBox txtKeyRun 
               Height          =   315
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   1380
               Width           =   2085
            End
            Begin VB.TextBox txtGroupRun 
               Height          =   315
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   4
               Top             =   690
               Width           =   2085
            End
            Begin VB.Label Label4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "ֵ��"
               Height          =   255
               Left            =   660
               TabIndex        =   9
               Top             =   3360
               Width           =   735
            End
            Begin VB.Label Label3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "������"
               Height          =   255
               Left            =   480
               TabIndex        =   7
               Top             =   2160
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "�ؼ��֣�"
               Height          =   255
               Left            =   300
               TabIndex        =   5
               Top             =   1440
               Width           =   735
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "�飺"
               Height          =   225
               Left            =   660
               TabIndex        =   3
               Top             =   750
               Width           =   735
            End
         End
      End
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "У���Ų�������Ϣ"
      Height          =   270
      Left            =   6420
      TabIndex        =   83
      Top             =   1110
      Width           =   1800
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'** �ļ�����frmOption.frm
'** ��  Ȩ��CopyRight (c) 2008-2010 �人��������ϵͳ���޹�˾
'** �����ˣ�yangshuai
'** ��  �䣺shuaigoplay@live.cn
'** ��  �ڣ�2009-2-27
'** �޸��ˣ�
'** ��  �ڣ�
'** ��  ����ϵͳ����
'** ��  ����1.0
'******************************************************************************

Dim sqlCtrl As String
Dim sqlRun As String
Dim sqlTpmsCode As String


Option Explicit
'�޸�TPMS��������ʼλ����Ϣ
Private Sub btMTOCModi_Click()
On Error GoTo Err
    If txtMtocStartIndex.text = "" Then
        MsgBox "TPMS��������ʼλ�ò���Ϊ��!"
        txtMtocStartIndex.SetFocus
        Exit Sub
    End If
    
    If txtMTOCLen.text = "" Then
        MsgBox "TPMS�����볤����Ϊ��!"
        txtMTOCLen.SetFocus
        Exit Sub
    End If

    Call updateRunParam(txtMtocStartIndex.text, "TPMSCode", "MTOCStartIndex")
    Call updateRunParam(txtMTOCLen.text, "TPMSCode", "TPMSCodeLen")
    
    mTOCStartIndex = txtMtocStartIndex.text
    tPMSCodeLen = txtMTOCLen.text
    
    MsgBox "TPMS��������ʼλ����Ϣ�޸ĳɹ�!"
    
    Exit Sub
Err:
    LogWritter "�޸�TPMS��������ʼλ����Ϣʱʧ�ܣ�����:" & Err.Description
    MsgBox "TPMS��������ʼλ����Ϣ�޸�ʧ��!" & Err.Description
End Sub

'����TPMS������
Private Sub btTPMSAdd_Click()
On Error GoTo Err
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    
    Set rs = cnn.Execute("select ""TPMSCode"" from ""T_TPMSCodeList"" where Upper(""TPMSCode"")='" & StrConv(txtTPMSCode.text, vbUpperCase) & "'")
    If Not rs.EOF Then
        MsgBox "��TPMS�������Ѵ���!"
        Exit Sub
    End If
    
    cnn.Execute ("insert into ""T_TPMSCodeList"" (""TPMSCode"") values ('" & txtTPMSCode.text & "')")
            
    cnn.Close
    Set cnn = Nothing
    
    showMSFlexGrid Me.MSFlexGrid3, DBCnnStr, sqlTpmsCode
    MsgBox "TPMS�����������ɹ�!"
    Exit Sub
Err:
    LogWritter "����TPMS������ʱʧ�ܣ�����:" & Err.Description
    MsgBox "TPMS����������ʧ��!" & Err.Description
End Sub

Private Sub btTPMSCancle_Click()
    Unload Me
End Sub
'ɾ��TPMS������
Private Sub btTPMSDel_Click()
On Error GoTo Err
    Dim msgR As Integer
    msgR = MsgBox("�Ƿ�ɾ��TPMS������" & txtTPMSCode.text & "��", vbYesNo, "ϵͳ��ʾ")
    If msgR = 7 Then Exit Sub

    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    
    cnn.Execute ("delete from ""T_TPMSCodeList"" where ""ID""=" & txtTPMSID.text & "")
            
    cnn.Close
    Set cnn = Nothing
    
    showMSFlexGrid Me.MSFlexGrid3, DBCnnStr, sqlTpmsCode
    MsgBox "TPMS������ɾ���ɹ�!"
    Exit Sub
Err:
    LogWritter "ɾ��TPMS������ʱʧ�ܣ�����:" & Err.Description
    MsgBox "TPMS������ɾ��ʧ��!" & Err.Description
End Sub

'�޸�TPMS������
Private Sub btTPMSModi_Click()
On Error GoTo Err
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    
    Set rs = cnn.Execute("select ""TPMSCode"" from ""T_TPMSCodeList"" where Upper(""TPMSCode"")='" & StrConv(txtTPMSCode.text, vbUpperCase) & "'")
    If Not rs.EOF Then
        MsgBox "��TPMS�������Ѵ���!"
        Exit Sub
    End If
    
    cnn.Execute ("update ""T_TPMSCodeList"" set ""TPMSCode""='" & txtTPMSCode.text & "' where ""ID""=" & txtTPMSID.text & "")
            
    cnn.Close
    Set cnn = Nothing
    
    showMSFlexGrid Me.MSFlexGrid3, DBCnnStr, sqlTpmsCode
    MsgBox "TPMS�������޸ĳɹ�!"
    Exit Sub
Err:
    LogWritter "�޸�TPMS������ʱʧ�ܣ�����:" & Err.Description
    MsgBox "TPMS�������޸�ʧ��!" & Err.Description
End Sub

'�Ƿ�����Ų�����
Private Sub chkAllQueue_Click()
    If chkAllQueue.value = vbChecked Then
        isCheckAllQueue = True
        Call updateRunParam(1, "Queue", "CheckAllQueue")
    Else
        isCheckAllQueue = False
        Call updateRunParam(0, "Queue", "CheckAllQueue")
    End If
End Sub

'�Ƿ�ֻ��ӡ��Ͻ��ΪNG����ϵ���
Private Sub chkOnlyPrintNGWriteResult_Click()
    If chkOnlyPrintNGWriteResult.value = vbChecked Then
        isOnlyPrintNGWriteResult = True
        Call updateRunParam(1, "Print", "OnlyPrintNGWriteResult")
    Else
        isOnlyPrintNGWriteResult = False
        Call updateRunParam(0, "Print", "OnlyPrintNGWriteResult")
    End If
End Sub

'��ɨ��VIN��
Private Sub chkOnlyScanVINCode_Click()
    If chkOnlyScanVINCode.value = vbChecked Then
        isOnlyScanVINCode = True
        Call updateRunParam(1, "Queue", "OnlyScanVINCode")
    Else
        isOnlyScanVINCode = False
        FrmMain.MTOCCode = "InitMTOCCode"
        Call updateRunParam(0, "Queue", "OnlyScanVINCode")
    End If
End Sub
'�Ƿ�ֻ��ӡNG��������̣��ϸ�����̲���ӡ
Private Sub chkPrintNGFlow_Click()
    If chkPrintNGFlow.value = vbChecked Then
        isOnlyPrintNGFlow = True
        Call updateRunParam(1, "Print", "OnlyPrintNGFlow")
    Else
        isOnlyPrintNGFlow = False
        Call updateRunParam(0, "Print", "OnlyPrintNGFlow")
    End If
End Sub

'�޸ļ��ٶȷ�Χֵ
Private Sub cmdAcSpeedSave_Click()
On Error GoTo Err
    If txtAcSpeedMin.text = "" Then
        MsgBox "���������ٶ���Сֵ����Ϊ��!"
        txtAcSpeedMin.SetFocus
        Exit Sub
    End If
    
    If txtAcSpeedMax.text = "" Then
        MsgBox "���������ٶ����ֵ����Ϊ��!"
        txtAcSpeedMax.SetFocus
        Exit Sub
    End If

    Call updateRunParam(txtAcSpeedMin.text, "StandardValue", "AcSpeedMinValue")
    Call updateRunParam(txtAcSpeedMax.text, "StandardValue", "AcSpeedMaxValue")
    
    acSpeedMinValue = txtAcSpeedMin.text
    acSpeedMaxValue = txtAcSpeedMax.text
    
    MsgBox "���������ٶ�ֵ��Χ�޸ĳɹ�!"
    
    Exit Sub
Err:
    LogWritter "�޸Ĵ��������ٶ�ֵʱʧ�ܣ�����:" & Err.Description
    MsgBox "���������ٶ�ֵ��Χ�޸�ʧ��!" & Err.Description
End Sub
'�޸�ģʽ
Private Sub cmdMdlSave_Click()
    If txtMdl.text = "" Then
        MsgBox "������ģʽ����Ϊ��!"
        txtMdl.SetFocus
        Exit Sub
    End If
    
    Call updateRunParam(txtMdl.text, "StandardValue", "MdlValue")
    mdlValue = txtMdl.text
End Sub

'�޸�ѹ����Χֵ
Private Sub cmdPreSave_Click()
On Error GoTo Err
    If txtPreMin.text = "" Then
        MsgBox "������ѹ����Сֵ����Ϊ��!"
        txtPreMin.SetFocus
        Exit Sub
    End If
    
    If txtPreMax.text = "" Then
        MsgBox "������ѹ�����ֵ����Ϊ��!"
        txtPreMax.SetFocus
        Exit Sub
    End If

    Call updateRunParam(txtPreMin.text, "StandardValue", "PreMinValue")
    Call updateRunParam(txtPreMax.text, "StandardValue", "PreMaxValue")
    
    preMinValue = txtPreMin.text
    preMaxValue = txtPreMax.text
    
    MsgBox "������ѹ��ֵ��Χ�޸ĳɹ�!"
    
    Exit Sub
Err:
    LogWritter "�޸Ĵ�����ѹ��ֵʱʧ�ܣ�����:" & Err.Description
    MsgBox "������ѹ��ֵ��Χ�޸�ʧ��!" & Err.Description
End Sub
'�޸��¶ȷ�Χֵ
Private Sub cmdTempSave_Click()
On Error GoTo Err
    If txtTempMin.text = "" Then
        MsgBox "�������¶���Сֵ����Ϊ��!"
        txtTempMin.SetFocus
        Exit Sub
    End If
    
    If txtTempMax.text = "" Then
        MsgBox "�������¶����ֵ����Ϊ��!"
        txtTempMax.SetFocus
        Exit Sub
    End If

    Call updateRunParam(txtTempMin.text, "StandardValue", "TempMinValue")
    Call updateRunParam(txtTempMax.text, "StandardValue", "TempMaxValue")
    
    tempMinValue = txtTempMin.text
    tempMaxValue = txtTempMax.text
    
    MsgBox "�������¶�ֵ��Χ�޸ĳɹ�!"
    
    Exit Sub
Err:
    LogWritter "�޸Ĵ������¶�ֵʱʧ�ܣ�����:" & Err.Description
    MsgBox "�������¶�ֵ��Χ�޸�ʧ��!" & Err.Description
End Sub

Private Sub ComboCtrl_Click()
    sqlCtrl = "select ""ID"" as ""���"",""Group"" as ""��"",""Description"" as ""����"",""Key"" as ""�ؼ���"",""Value"" as ""ֵ"" from ""T_CtrlParam"" where ""Group""='" & Me.ComboCtrl.text & "'  order by ""ID"" "
    showMSFlexGrid Me.MSFlexGrid2, DBCnnStr, sqlCtrl
End Sub


Private Sub ComboRun_Click()
    sqlRun = "select ""ID"" as ""���"",""Group"" as ""��"",""Description"" as ""����"",""Key"" as ""�ؼ���"",""Value"" as ""ֵ"" from ""T_RunParam"" where ""Group""='" & Me.ComboRun.text & "' order by ""ID""  "
    showMSFlexGrid Me.MSFlexGrid1, DBCnnStr, sqlRun
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command10_Click()
    frmQueueInfo.Show
End Sub

Private Sub Command2_Click()
    On Error GoTo update_err
    updateParam "Run", Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
    showMSFlexGrid Me.MSFlexGrid1, DBCnnStr, sqlRun
    Exit Sub
update_err:
    MsgBox "�޸Ĵ��󣬴�����Ϣ��" & Err.Description
    
End Sub

Private Sub Command3_Click()
    On Error GoTo update_err
    updateParam "Ctrl", Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 0)
    showMSFlexGrid Me.MSFlexGrid2, DBCnnStr, sqlCtrl
    Exit Sub
update_err:
    MsgBox "�޸Ĵ��󣬴�����Ϣ��" & Err.Description
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

'�ֶ������Ų���������
Private Sub Command5_Click()
On Error GoTo Err
    Dim diaFlag As Integer

    If FrmMain.TestStateFlag <= 5 Then
        diaFlag = MsgBox("�������ڽ���̥ѹ��⣬���Ժ��������Ų�������Ϣ", vbOKOnly, "ϵͳ��ʾ")
        Exit Sub
    End If

    diaFlag = MsgBox("�Ƿ������Ų�������Ϣ?", vbYesNo, "ϵͳ��ʾ")
    If diaFlag = 7 Then
      Exit Sub
    End If

    If Not Ping(MES_IP) Then
        diaFlag = MsgBox("����MES������ʱʧ�ܣ���������״̬�Ƿ�ͨ!", vbOKOnly, "ϵͳ��ʾ")
        Exit Sub
    End If

    Dim objConn As Connection
    Dim objConnMES As Connection
    Dim objRs As Recordset
    Dim objTmpRs As Recordset
    Dim objRsMES As Recordset
    Dim strSQL As String

    '�ȶ�ȡMES�ϵ�����
    Set objConnMES = New Connection
    Set objRsMES = New Recordset
    objConnMES.ConnectionTimeout = 3
    DoEvents
    objConnMES.Open MESCnnStr
    If objConnMES.state <> adStateOpen Then
        diaFlag = MsgBox("����MES���ݿ�ʱʧ�ܣ�����Oracle�ͻ���������Ϣ�Ƿ���ȷ!", vbOKOnly, "ϵͳ��ʾ")
        Set objConnMES = Nothing
        Exit Sub
    End If
    LogWritter "�����ֶ�ͬ���Ų���������"
    strSQL = "select * from mesprd.IF_VEHICLE_TPMS_INFO where tpms_process=0 order by pa_off_seq asc"
    'strSQL = "update mesprd.IF_VEHICLE_TPMS_INFO set tpms_process=0 where pa_off_seq>=18452"
    objRsMES.Open strSQL, objConnMES, adOpenKeyset, adLockOptimistic

    '�򿪱������ݿ�����
    Set objConn = New Connection
    Set objRs = New Recordset
    objConn.ConnectionTimeout = 2
    objConn.Open DBCnnStr

    strSQL = "select * from vinlist"
    objRs.Open strSQL, objConn, adOpenStatic, adLockOptimistic
    DoEvents
    Set objTmpRs = New Recordset
    Do While Not objRsMES.EOF              '---���������
        strSQL = "select * from vinlist where vin='" & objRsMES("vin") & "'"
        objTmpRs.Open strSQL, objConn, adOpenStatic, adLockOptimistic
        If objTmpRs.EOF Then
            objRs.AddNew
            objRs("vin") = objRsMES("vin")
            objRs!mtoc = objRsMES!mtoc
            objRs!pa_off_seq = objRsMES!pa_off_seq
            objRs!pa_off_time = objRsMES!pa_off_time
            objRs!createtime = Now()
            objRs.Update
        Else
            objTmpRs!mtoc = objRsMES!mtoc
            objTmpRs!pa_off_seq = objRsMES!pa_off_seq
            objTmpRs!pa_off_time = objRsMES!pa_off_time
            objTmpRs!createtime = Now()
            objTmpRs.Update
        End If

        '����MESϵͳ�����ر�ʶ
        strSQL = "update mesprd.IF_VEHICLE_TPMS_INFO set tpms_process=1 where vin='" & objRsMES("vin") & "'"
        objConnMES.Execute strSQL

        objRsMES.MoveNext
        objTmpRs.Close
    Loop
    objRs.Close
    objRsMES.Close
    objConn.Close
    objConnMES.Close
    Set objRs = Nothing
    Set objTmpRs = Nothing
    Set objRsMES = Nothing
    Set objConn = Nothing
    Set objConnMES = Nothing

    LogWritter "�Ų���������ͬ�����"
    diaFlag = MsgBox("�Ų������������سɹ�!", vbOKOnly, "ϵͳ��ʾ")
    Exit Sub
Err:
    LogWritter Err.Description
    diaFlag = MsgBox(Err.Description, vbOKOnly, "ϵͳ��ʾ")
End Sub

Private Sub Command6_Click()
    On Error GoTo Err
    Dim objConn As Connection
    Dim objRs As Recordset
    Dim strSQL As String
        
    If Text1.text = Text2.text And Text1.text <> "" Then
        
        '�򿪱������ݿ�����
        Set objConn = New Connection
        Set objRs = New Recordset
        objConn.ConnectionTimeout = 2
        objConn.Open DBCnnStr
        
        strSQL = "UPDATE ""T_Psw"" SET ""psw"" = '" & Text1.text & "'"
        objRs.Open strSQL, objConn, adOpenStatic, adLockOptimistic
        objConn.Close
        Set objRs = Nothing
        Set objConn = Nothing
        MsgBox "���������޸ĳɹ�"
        LogWritter "���������޸ĳɹ�"
    
    Else
        MsgBox "�������벻��Ϊ��"
    End If
    Exit Sub
Err:
    LogWritter "�޸�������̳���"
End Sub

Private Sub Command7_Click()
    If txtVIN.text = "" Then
        MsgBox "��ӡVIN����Ϊ��!"
        txtVIN.SetFocus
        Exit Sub
    End If

    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select ""VIN"" from ""T_Result"" where ""VIN""='" & txtVIN.text & "'")
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        cnn.Close
        Set cnn = Nothing
        MsgBox "ϵͳ�в����ڸó�����ؼ����Ϣ!"
        Exit Sub
    End If

    printErrCodeByVIN (txtVIN.text)
End Sub

Private Sub Form_Load()
    WindowsXPC1.InitSubClassing
    Me.SSTab1.Tab = 0
    Me.SSTab1.TabVisible(3) = False

    sqlCtrl = "Select ""ID"" as ""���"",""Group"" as ""��"",""Description"" as ""����"",""Key"" as ""�ؼ���"",""Value"" as ""ֵ"" from ""T_CtrlParam"" order by ""ID"" "
    sqlRun = "Select ""ID"" as ""���"",""Group"" as ""��"",""Description"" as ""����"",""Key"" as ""�ؼ���"",""Value"" as ""ֵ"" from ""T_RunParam"" order by ""ID"" "
    sqlTpmsCode = "select ""ID"",""ID"" as ""���"",""TPMSCode"" as ""TPMS������"" from ""T_TPMSCodeList"" order by ""ID"""
    '���������
    loadCombo Me.ComboRun, "T_RunParam"
    showMSFlexGrid Me.MSFlexGrid1, DBCnnStr, sqlRun
    loadCombo Me.ComboCtrl, "T_CtrlParam"
    showMSFlexGrid Me.MSFlexGrid2, DBCnnStr, sqlCtrl
    showMSFlexGrid Me.MSFlexGrid3, DBCnnStr, sqlTpmsCode
    Me.MSFlexGrid3.ColWidth(1) = 800
    
    If isCheckAllQueue Then
        chkAllQueue.value = 1
    Else
        chkAllQueue.value = 0
    End If
    If isOnlyScanVINCode Then
        chkOnlyScanVINCode.value = 1
    Else
        chkOnlyScanVINCode.value = 0
    End If
    If isOnlyPrintNGWriteResult Then
        chkOnlyPrintNGWriteResult.value = 1
    Else
        chkOnlyPrintNGWriteResult.value = 0
    End If
    If isOnlyPrintNGFlow Then
        chkPrintNGFlow.value = 1
    Else
        chkPrintNGFlow.value = 0
    End If
    
    txtMdl.text = mdlValue
    txtPreMin.text = preMinValue
    txtPreMax.text = preMaxValue
    txtTempMin.text = tempMinValue
    txtTempMax.text = tempMaxValue
    txtAcSpeedMin.text = acSpeedMinValue
    txtAcSpeedMax.text = acSpeedMaxValue
    txtMtocStartIndex.text = mTOCStartIndex
    txtMTOCLen.text = tPMSCodeLen
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

'******************************************************************************
'** �� �� ����showMSFlexGrid
'** ��    �룺
'** ��    ����
'** ������������ʾ���
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************

Public Sub showMSFlexGrid(msFG As MSFlexGrid, CnnStr As String, sql As String)
On Error GoTo Err_ShowGrid
    msFG.Clear
    If sql = "" Then
        Exit Sub
    End If
    Dim cnn As New ADODB.Connection
    Dim rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim i As Integer, J As Integer
    
    cnn.Open CnnStr
    rs.Open sql, cnn, 1, 3
    
    With msFG
        .Visible = True
        .cols = rs.Fields.Count
        .Rows = rs.RecordCount + 11
        .FillStyle = 1
        '.CellAlignment = flexAlignLeftCenter
        For i = 0 To rs.Fields.Count - 1
            .TextMatrix(0, i) = rs.Fields(i).Name
        Next
        J = 1
        Do While Not rs.EOF
            For i = 0 To rs.Fields.Count - 1
                If IsNull(rs(i)) Then
                    .TextMatrix(J, i) = ""
                Else
                    .TextMatrix(J, i) = rs(i)
                    
                End If
            Next
            rs.MoveNext
            J = J + 1
        Loop
    End With
    Call setColWidth(msFG, rs.Fields.Count)  '�����п�������̿��Ը����Լ���Ҫ����
    rs.Close
    Set rs = Nothing
    cnn.Close
    Exit Sub
Err_ShowGrid:
    MsgBox "��ʾ���ݳ���������Ϣ��" & Err.Description
End Sub
Private Sub setColWidth(msFG As MSFlexGrid, cols As Integer)
With msFG
    Dim i As Integer
    .ColWidth(0) = 0
    .ColWidth(1) = 800
    For i = 2 To cols - 1 'Ϊÿ���е��н�������
        .ColWidth(i) = 1500 '�еĿ��,�Ժ��Լ�����
    Next

End With
End Sub

Private Sub MSFlexGrid1_Click()
    On Error Resume Next
    txtGroupRun.text = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    txtDescriptionRun.text = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    txtKeyRun.text = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)
    txtValueRun.text = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4)
    
End Sub

Private Sub MSFlexGrid2_Click()
    On Error Resume Next
    txtGroupCtrl.text = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 1)
    txtDescriptionCtrl.text = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 2)
    txtKeyCtrl.text = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 3)
    txtValueCtrl.text = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 4)
    'showMSFlexGrid Me.MSFlexGrid2, DBCnnStr, sqlCtrl
End Sub



'******************************************************************************
'** �� �� ����updateParam
'** ��    �룺
'** ��    ����
'** �����������޸�����
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Public Sub updateParam(typeStr As String, id As Long)
    Dim cnn As New ADODB.Connection
    Dim tableName As String
    Dim textName As String
    tableName = Chr(34) & "T_" & typeStr & "Param" & Chr(34)
    textName = "txtValue" & typeStr
    cnn.Open DBCnnStr
    cnn.Execute "update " & tableName & " set ""Value""='" & Me.Controls(textName).text & "' where ""ID""=" & id
    cnn.Close
    Set cnn = Nothing
End Sub


'******************************************************************************
'** �� �� ����loadCombo����Combo�ؼ�����
'** ��    �룺
'** ��    ����
'** �����������޸�����
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Private Sub loadCombo(combo As ComboBox, tableName As String)
    On Error GoTo loadCombo_err
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select ""Group"" from """ & tableName & """ group by ""Group""  ")
    combo.Clear
    Do While Not rs.EOF
        combo.AddItem rs(0).value
        rs.MoveNext
    Loop
    cnn.Close
    Exit Sub
loadCombo_err:
    MsgBox "���ش��󣡴�����Ϣ��" & Err.Description
    
End Sub
Private Sub MSFlexGrid3_Click()
    On Error Resume Next
    txtTPMSID.text = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 1)
    txtTPMSCode.text = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 2)
End Sub
Public Function readRunParam(key As String, group As String) As String
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select ""Value"" from ""T_RunParam"" where ""Key""='" & key & "' and ""Group""='" & group & "'")
    readRunParam = rs("Value")
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Function
Public Function updateRunParam(value As String, group As String, key As String)
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    cnn.Open DBCnnStr
    cnn.Execute ("update ""T_RunParam"" set ""Value""='" & value & "'  where ""Key""='" & key & "' and ""Group""='" & group & "'")
    cnn.Close
    Set cnn = Nothing
End Function
