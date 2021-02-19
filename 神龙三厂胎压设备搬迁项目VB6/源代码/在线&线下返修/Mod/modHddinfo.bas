Attribute VB_Name = "Modhddinfo"
Option Explicit
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Public Type LARGE_INTEGER
    Lowpart As Long
    Highpart As Long
End Type
Public Result As Double
Public Const SIZE_KB As Double = 1024
Public Const SIZE_MB As Double = 1024 * SIZE_KB
Public Const SIZE_GB As Double = 1024 * SIZE_MB
Public Const SIZE_TB As Double = 1024 * SIZE_GB
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const DRIVE_UNKNOWN = 0
Public Const DRIVE_NOTEXIST = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_CDROM = 5

Public Const FILE_CASE_SENSITIVE_SEARCH = &H1
Public Const FILE_CASE_PRESERVED_NAMES = &H2
Public Const FILE_UNICODE_ON_DISK = &H4
Public Const FILE_PERSISTENT_ACLS = &H8
Public Const FILE_FILE_COMPRESSION = &H10
Public Const FILE_VOLUME_QUOTAS = &H20
Public Const FILE_SUPPORTS_SPARSE_FILES = &H40
Public Const FILE_SUPPORTS_REPARSE_POINTS = &H80
Public Const FILE_SUPPORTS_REMOTE_STORAGE = &H100
Public Const FILE_VOLUME_IS_COMPRESSED = &H8000
Public Const FILE_SUPPORTS_OBJECT_IDS = &H10000
Public Const FILE_SUPPORTS_ENCRYPTION = &H20000
Public Const FILE_NAMED_STREAMS = &H40000

Public Const FS_CASE_IS_PRESERVED = FILE_CASE_PRESERVED_NAMES
Public Const FS_CASE_SENSITIVE = FILE_CASE_SENSITIVE_SEARCH
Public Const FS_UNICODE_STORED_ON_DISK = FILE_UNICODE_ON_DISK
Public Const FS_PERSISTENT_ACLS = FILE_PERSISTENT_ACLS
Public Const FS_VOL_IS_COMPRESSED = FILE_VOLUME_IS_COMPRESSED
Public Const FS_FILE_COMPRESSION = FILE_FILE_COMPRESSION
Public Const FS_FILE_ENCRYPTION = FILE_SUPPORTS_ENCRYPTION

Public sDriveNames As String
Public lBuffer As Long
Public lReturn As Long
Public nLoopCtr As Integer
Public nOffset As Integer
Public sTempStr As String

Public Root As String
Public Volume_Name As String
Public Serial_Number As Long
Public Max_Component_Length As Long
Public File_System_Flags As Long
Public File_System_Name As String
Public Pos As Integer
Public Dbl_Total As Double
Public Dbl_Free As Double

Public lSectorsPerCluster As Long
Public lBytesPerSector As Long
Public lFreeClusters As Long
Public lTotalClusters As Long
Public sDrive As String
Public Function GetHDDState(HDDID As String, FreeHDD As Long) As Integer
Dim TempHDDID As String
Dim Bytes_Avail As LARGE_INTEGER
Dim Bytes_Total As LARGE_INTEGER
Dim Bytes_Free As LARGE_INTEGER
TempHDDID = Trim(UCase(HDDID) & ":\")
On Error Resume Next
If GetVolumeInformation(TempHDDID, Volume_Name, Len(Volume_Name), Serial_Number, Max_Component_Length, File_System_Flags, File_System_Name, Len(File_System_Name)) = 0 Then
    GetHDDState = 2         '"�̷�����"
    Exit Function
End If
    GetDiskFreeSpaceEx TempHDDID, Bytes_Avail, Bytes_Total, Bytes_Free
Dbl_Free = LargeIntegerToDouble(Bytes_Free.Lowpart, Bytes_Free.Highpart)
    If SizeString(Dbl_Free) < FreeHDD Then
        GetHDDState = 0             'UCase(HDDID) & "�̿ռ�����"
    Else
        GetHDDState = 1             'UCase(HDDID) & "�̿ռ�����"
    End If
End Function
Private Function LargeIntegerToDouble(Low_Part As Long, High_Part As Long) As Double
Result = High_Part
If High_Part < 0 Then Result = Result + 2 ^ 32
    Result = Result * 2 ^ 32
    Result = Result + Low_Part
If Low_Part < 0 Then Result = Result + 2 ^ 32
    LargeIntegerToDouble = Result
End Function
Private Function SizeString(ByVal Num_Bytes As Double) As String
        SizeString = Format$(Num_Bytes / SIZE_MB, "0.00")
End Function
