VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJobGetVIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IJob

Private vinState As Boolean
Private WithEvents comm As MSComm
Attribute comm.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set comm = FrmMain.MSCommVIN
End Sub



Private Sub comm_OnComm()
    
End Sub

Private Function IJob_JobToDo() As Boolean
    
End Function



