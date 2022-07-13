VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------------------------
'  Copyright ©2001 Mike G --> IvbNET.COM
'  All Rights Reserved, http://www.ivbnet.com
'  EMAIL : webmaster@ivbnet.com
' ------------------------------------------------------------
'  You are free to use this code within your own applications,
'  but you are forbidden from selling or distributing this
'  source code without prior written consent.
' ------------------------------------------------------------


Function FirstOfMonth(Optional dteDate As Date) As Date
      ' This function calculates the first day of a month, given a date.
      ' If no date is passed in, the function uses the current date.
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      ' Find the first day of this month.
      FirstOfMonth = DateSerial(Year(dteDate), Month(dteDate), 1)
End Function

Function LastOfMonth(Optional dteDate As Date) As Date
      ' This function calculates the last day of a month, given a date.
      ' If no date is passed in, the function uses the current date.
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      ' Find the first day of the next month, then subtract one day.
      LastOfMonth = DateSerial(Year(dteDate), Month(dteDate) + 1, 1) - 1
End Function

Function DayOfYear(Optional dteDate As Date) As Long
      ' If dteDate argument has not been passed, dteDate is
      ' initialized to 0 (or December 30, 1899, the date
      ' equivalent of 0).
      If CLng(dteDate) = 0 Then
            ' Use today's date.
            dteDate = Date
      End If
      ' Calculate the number of days that have passed since
      ' December 31 of the previous year.
      DayOfYear = Abs(DateDiff("d", dteDate, _
                  DateSerial(Year(dteDate) - 1, 12, 31)))
End Function

Function IsWorkday(Optional dteDate As Date) As Boolean
      ' This function determines whether a date
      ' falls on a weekday.
      ' If no date passed in, use today's date.
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      ' Determine where in week the date falls.
      Select Case Weekday(dteDate)
            Case vbMonday To vbFriday
                  IsWorkday = True
            Case Else
                  IsWorkday = False
      End Select
End Function

Function Anniversary(dteDate As Date) As Date
      ' This function finds the next anniversary of a date.
      ' If the date has already passed for this year, it returns
      ' the date on which the anniversary occurs in the following year.

      Dim dteThisYear As Date

      ' Find corresponding date this year.
      dteThisYear = DateSerial(Year(Date), Month(dteDate), Day(dteDate))
      ' Determine whether it's already passed.
      If dteThisYear < Date Then
            Anniversary = DateAdd("yyyy", 1, dteThisYear)
      Else
            Anniversary = dteThisYear
      End If
End Function

Function CalcAge(dteBirthdate As Date) As Long

      Dim lngAge As Long

      ' Make sure passed-in value is a date.
      If Not IsDate(dteBirthdate) Then
            dteBirthdate = Date
      End If

      ' Make sure birthdate is not in the future.
      ' If it is, use today's date.
      If dteBirthdate > Date Then
            dteBirthdate = Date
      End If

      ' Calculate the difference in years between today and birthdate.
      lngAge = DateDiff("yyyy", dteBirthdate, Date)
      ' If birthdate has not occurred this year, subtract 1 from age.
      If DateSerial(Year(Date), Month(dteBirthdate), Day(dteBirthdate)) > Date Then
            lngAge = lngAge - 1
      End If
      CalcAge = lngAge
End Function

'Function GetElapsedTime(dteStart As Date, _
'                        dteEnd As Date, _
'                        lngFormat As opgTimeInterval) As String'

      ' Formats elapsed time as seconds, minutes:seconds, hours:minutes:seconds,
      ' or days, hours, minutes, seconds.

'      Dim dblInterval     As Double
'      Dim lngTotalHours   As Long
'      Dim lngTotalMins    As Long
'      Dim lngTotalSecs    As Long
'      Dim lngDays         As Long
'      Dim lngHours        As Long
'      Dim lngMins         As Long
'      Dim lngSecs         As Long
'      Dim strElapsed      As String

'      dblInterval = Abs(dteEnd - dteStart)

'      lngDays = Int(CSng(dblInterval))
'      lngTotalHours = Int(CSng(dblInterval * 24))
'      lngTotalMins = Int(CSng(dblInterval * 1440))
'      lngTotalSecs = Int(CSng(dblInterval * 86400))

      ' Determine fractional times.
'      lngHours = lngTotalHours Mod 24
'      lngMins = lngTotalMins Mod 60
'      lngSecs = lngTotalSecs Mod 60

'      Select Case lngFormat
                  ' Determine total number of seconds.
'            Case opgTimeInterval.SEC
'                  strElapsed = lngTotalSecs & " Seconds"
                  ' Determine number of minutes plus number of seconds comprising fractional minute.
'            Case opgTimeInterval.MIN_SEC
'                  strElapsed = lngTotalMins & ":" & Format$(dblInterval, "ss") _
                              & " Minutes:Seconds"
                  ' Determine number of hours; express fraction as minutes and seconds.
'            Case opgTimeInterval.HOUR_MIN_SEC
'                  strElapsed = lngTotalHours & ":" & Format$(dblInterval, "nn") _
'                              & ":" & Format$(dblInterval, "ss") & " Hours:Minutes:Seconds"
                  ' Determine number of days; express fraction as hours, minutes, seconds.
'            Case opgTimeInterval.DAY_HOUR_MIN_SEC
'                  strElapsed = lngDays & " days " & lngHours _
'                              & " Hours " & lngMins & " Minutes " & _
                              lngSecs & " Seconds"
'      End Select

'      GetElapsedTime = strElapsed
'End Function

Function FirstOfNextMonth(Optional dteDate As Date) As Date
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      FirstOfNextMonth = DateSerial(Year(dteDate), Month(dteDate) + 1, 1)
End Function

Function LastOfNextMonth(Optional dteDate As Date) As Date
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      LastOfNextMonth = DateSerial(Year(dteDate), Month(dteDate) + 2, 0)
End Function

Function FirstOfPreviousMonth(Optional dteDate As Date) As Date
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      FirstOfPreviousMonth = DateSerial(Year(dteDate), Month(dteDate) - 1, 1)
End Function

Function LastOfPreviousMonth(Optional dteDate As Date) As Date
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      LastOfPreviousMonth = DateSerial(Year(dteDate), Month(dteDate), 0)
End Function

Function FirstOfQuarter(Optional dteDate As Date) As Date
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      FirstOfQuarter = DateSerial(Year(dteDate), Int((Month(dteDate) - 1) / 3) * 3 + 1, 1)
End Function

Function LastOfQuarter(Optional dteDate As Date) As Date
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      LastOfQuarter = DateSerial(Year(Date), Int((Month(Date) - 1) / 3) * 3 + 4, 0)
End Function

Function FirstOfWeek(Optional dteDate As Date) As Date
      ' Assuming Sunday is first day of week.
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      FirstOfWeek = dteDate - Weekday(dteDate) + 1
End Function

Function LastOfWeek(Optional dteDate As Date) As Date
      ' Assuming Sunday is first day of week.
      If CLng(dteDate) = 0 Then
            dteDate = Date
      End If
      LastOfWeek = dteDate - Weekday(dteDate) + 7
End Function

Function Age(Bdate, DateToday) As Integer
      ' Returns the Age in years between 2 dates
      ' Doesn't handle negative date ranges i.e. Bdate > DateToday
      If Month(DateToday) < Month(Bdate) Or (Month(DateToday) = Month(Bdate) And Day(DateToday) < Day(Bdate)) Then
            Age = Year(DateToday) - Year(Bdate) - 1
      Else
            Age = Year(DateToday) - Year(Bdate)
      End If
End Function

Function DaysInMonth(D As Variant) As Variant
      ' Returns the number of days in a month.
      ' Requires a date argument, since February can change if it's a leap year
     
      If VarType(D) <> 7 Then
            DaysInMonth = Null
      Else
            Select Case Month(D)
                  Case 2
                        If LeapYear(Year(D)) Then
                              DaysInMonth = 29
                        Else
                              DaysInMonth = 28
                        End If
                  Case 4, 6, 9, 11
                        DaysInMonth = 30
                  Case 1, 3, 5, 7, 8, 10, 12
                        DaysInMonth = 31
            End Select
      End If
End Function

Function DaysInMonth2(D As Variant) As Variant
      ' Returns the number of days in a month
      ' Requires a date argument, since February can change if it's a leap year
      ' Lets Access figure it out
      If VarType(D) <> 7 Then
            DaysInMonth2 = Null
      Else
            DaysInMonth2 = DateSerial(Year(D), Month(D) + 1, 1) - DateSerial(Year(D), Month(D), 1)
      End If
End Function

Function EndOfMonth(D As Variant) As Variant
      ' Returns the date representing the last day of the current month.
      ' Arguments:
      ' D            = Date
      EndOfMonth = DateSerial(Year(D), Month(D) + 1, 0)
End Function

Function EndOfWeek(D As Variant, _
                  Optional FirstWeekday As Integer) As Variant
        ' Returns the date representing the last day of the current week.
      ' Arguments:
      ' D            = Date
      ' FirstWeekday = (Optional argument) Integer that represents the first
      ' day of the week (e.g., 1=Sun..7=Sat).
      '
      If IsMissing(FirstWeekday) Then                 'Sunday is the assumed first day of week.
            EndOfWeek = D - Weekday(D) + 7
      Else
            EndOfWeek = D - Weekday(D, FirstWeekday) + 7
      End If
End Function


Function LastBusDay(D As Variant) As Variant
      ' Returns the date of the last business day (Mon-Fri) in a month
      '
      Dim D2 As Variant
      If VarType(D) <> 7 Then
            LastBusDay = Null
      Else
            D2 = DateSerial(Year(D), Month(D) + 1, 0)
            Do While Weekday(D2) = 1 Or Weekday(D2) = 7
                  D2 = D2 - 1
            Loop
            LastBusDay = D2
      End If
End Function

Function LeapYear(YYYY As Integer) As Integer
      ' Leap Year from standard rules
      ' YYYY: 4-digit year
      LeapYear = YYYY Mod 4 = 0 And (YYYY Mod 100 <> 0 Or YYYY Mod 400 = 0)
End Function

Function LeapYear2(YYYY As Integer) As Integer
      ' Leap Year letting Access figure out the rules
      ' YYYY: 4-digit year
      LeapYear2 = Month(DateSerial(YYYY, 2, 29)) = 2
End Function

Function NextDay(D As Variant, DayCode As Integer) As Variant
      ' Returns the date of the next DayCode (1=Sun ... 7=Sat) after the
      ' date D.  e.g.  NextDay(#5/12/94#,6) returns the date of the next
      ' Friday after 5/12/94.
      NextDay = D - Weekday(D) + DayCode + IIf(Weekday(D) < DayCode, 0, 7)
End Function

Function NextDay1(D As Variant, DayCode As Integer) As Variant
      ' Returns the date of the next DayCode (1=Sun ... 7=Sat) on or after the
      ' date D.
      ' e.g.  NextDay1(#5/12/94#,6) returns the date of the next Friday after 5/12/94,
      ' or, if 5/12/94 is a Friday, then returns that date (5/12/94).
        NextDay1 = D - Weekday(D) + DayCode + IIf(Weekday(D) <= DayCode, 0, 7)
End Function

Function PriorDay(D As Variant, DayCode As Integer) As Variant
      ' Returns the date of the last DayCode (1=Sun ... 7=Sat) before the
      ' date D.  e.g.  PriorDay(#5/12/94#,6) returns the date of the
      ' Friday prior to 5/12/94.
      PriorDay = D - Weekday(D) + DayCode - IIf(Weekday(D) > DayCode, 0, 7)
End Function

Function PriorDay1(D As Variant, DayCode As Integer) As Variant
      ' Returns the date of the last DayCode (1=Sun ... 7=Sat) on or
      ' before the date D.
      ' e.g.  PriorDay1(#5/12/94#,6) returns the date of the Friday prior to 5/12/94,
      ' or if 5/12/94 is a Friday, then returns that date (5/12/94).
  
      PriorDay1 = D - Weekday(D) + DayCode - IIf(Weekday(D) >= DayCode, 0, 7)
End Function

Function StartOfMonth(D As Variant) As Variant
      ' Returns the date representing the first day of the current month.
      ' Arguments:
      ' D            = Date
      StartOfMonth = DateSerial(Year(D), Month(D), 1)
End Function

Function StartOfWeek(D As Variant, _
                     Optional FirstWeekday As Integer) As Variant      '
      ' Returns the date representing the first day of the current week.
      '
      ' Arguments:
      ' D            = Date
      ' FirstWeekday = (Optional argument) Integer that represents the first
      ' day of the week (e.g., 1=Sun..7=Sat).
      '
      If IsMissing(FirstWeekday) Then                 'Sunday is the assumed first day of week.
            StartOfWeek = D - Weekday(D) + 1
      Else
            StartOfWeek = D - Weekday(D, FirstWeekday) + 1
      End If
End Function

Public Function IsLeapYear(DateIn As Date) As Boolean

      If IsDate("29/02/" & Format(DateIn, "yyyy")) = True Then
            IsLeapYear = True
      End If

End Function






