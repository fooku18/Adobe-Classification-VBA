VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalToUTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
'Purpose     :  Converts local time to GMT.
'Inputs      :  dtLocalDate                 The local data time to return as GMT.
'Outputs     :  Returns the local time in GMT.
'Author      :  Andrew Baker
'Date        :  13/11/2002 10:16
'Notes       :
'Revisions   :

Public Function ConvertLocalToGMT(dtLocalDate As Date) As Date
    Dim lSecsDiff As Long

    'Get the GMT time diff
    lSecsDiff = GetLocalToGMTDifference()
    'Return the time in GMT
    ConvertLocalToGMT = DateAdd("s", -lSecsDiff, dtLocalDate)
End Function

'Purpose     :  Converts GMT time to local time.
'Inputs      :  dtLocalDate                 The GMT data time to return as local time.
'Outputs     :  Returns GMT as local time.
'Author      :  Andrew Baker
'Date        :  13/11/2002 10:16
'Notes       :
'Revisions   :

Public Function ConvertGMTToLocal(gmtTime As Date) As Date
    Dim Differerence As Long

    Differerence = GetLocalToGMTDifference()
    ConvertGMTToLocal = DateAdd("s", Differerence, gmtTime)
End Function

'Purpose     :  Returns the time lDiff between local and GMT (secs).
'Inputs      :  dtLocalDate                 The local data time to return as GMT.
'Outputs     :  Returns the local time in GMT.
'Author      :  Andrew Baker
'Date        :  13/11/2002 10:16
'Notes       :  A positive number indicates your ahead of GMT.
'Revisions   :

Public Function GetLocalToGMTDifference() As Long
    Const TIME_ZONE_ID_INVALID& = &HFFFFFFFF
    Const TIME_ZONE_ID_STANDARD& = 1
    Const TIME_ZONE_ID_UNKNOWN& = 0
    Const TIME_ZONE_ID_DAYLIGHT& = 2

    Dim tTimeZoneInf As TIME_ZONE_INFORMATION
    Dim lRet As Long
    Dim lDiff As Long

    'Get time zone info
    lRet = GetTimeZoneInformation(tTimeZoneInf)

    'Convert diff to secs
    lDiff = -tTimeZoneInf.Bias * 60
    GetLocalToGMTDifference = lDiff

    'Check if we are in daylight saving time.
    If lRet = TIME_ZONE_ID_DAYLIGHT& Then
        'In daylight savings, apply the bias
        If tTimeZoneInf.DaylightDate.wMonth <> 0 Then
            'if tTimeZoneInf.DaylightDate.wMonth = 0 then the daylight
            'saving time change doesn't occur
            GetLocalToGMTDifference = lDiff - tTimeZoneInf.DaylightBias * 60
        End If
    End If
End Function
