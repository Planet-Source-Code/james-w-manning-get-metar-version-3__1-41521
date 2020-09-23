Attribute VB_Name = "TimeZone"
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(32) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(32) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type
Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Function GetLocalTZ(Optional ByRef strTZName As String) As Long
    Dim objTimeZone As TIME_ZONE_INFORMATION
    Dim lngResult As Long
    Dim i As Long
    lngResult = GetTimeZoneInformation&(objTimeZone)


    Select Case lngResult
        Case 0, 1 'use standard time
            GetLocalTZ = -(objTimeZone.Bias / 60)
    
    
            For i = 0 To 31
                If objTimeZone.StandardName(i) = 0 Then Exit For
                strTZName = strTZName & Chr(objTimeZone.StandardName(i))
            Next
        Case 2 'use daylight savings time
            GetLocalTZ = -(objTimeZone.DaylightBias / 60)
    
    
            For i = 0 To 31
                If objTimeZone.DaylightName(i) = 0 Then Exit For
                strTZName = strTZName & Chr(objTimeZone.DaylightName(i))
            Next
    End Select
End Function

