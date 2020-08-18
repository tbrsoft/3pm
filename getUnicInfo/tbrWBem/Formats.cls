VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Formats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior = 0   'vbNone
'  MTSTransactionMode = 0   'NotAnMTSObject
'End
'Attribute VB_Name = "Formats"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = True
Option Explicit

Public Enum LocaleInfo
   LOCALE_IDIGITS& = &H11
   LOCALE_ILZERO& = &H12
   LOCALE_STIMEFORMAT& = &H1003
   LOCALE_STIME& = &H1E
   LOCALE_STHOUSAND& = &HF
   LOCALE_SSHORTDATE& = &H1F
   LOCALE_SPOSITIVESIGN& = &H50
   LOCALE_SNEGATIVESIGN& = &H51
   LOCALE_SNATIVELANGNAME& = &H4
   LOCALE_SMONGROUPING& = &H18
   LOCALE_SMONTHOUSANDSEP& = &H17
   LOCALE_SMONTHNAME9& = &H40
   LOCALE_SMONTHNAME8& = &H3F
   LOCALE_SMONTHNAME7& = &H3E
   LOCALE_SMONTHNAME6& = &H3D
   LOCALE_SMONTHNAME5& = &H3C
   LOCALE_SMONTHNAME4& = &H3B
   LOCALE_SMONTHNAME3& = &H3A
   LOCALE_SMONTHNAME2& = &H39
   LOCALE_SMONTHNAME12& = &H43
   LOCALE_SMONTHNAME11& = &H42
   LOCALE_SMONTHNAME10& = &H41
   LOCALE_SMONTHNAME1& = &H38
   LOCALE_SMONDECIMALSEP& = &H16
   LOCALE_SLONGDATE& = &H20
   LOCALE_SLANGUAGE& = &H2
   LOCALE_SGROUPING& = &H10
   LOCALE_SDECIMAL& = &HE
   LOCALE_SDAYNAME7& = &H30
   LOCALE_SDAYNAME6& = &H2F
   LOCALE_SDAYNAME5& = &H2E
   LOCALE_SDAYNAME4& = &H2D
   LOCALE_SDAYNAME3& = &H2C
   LOCALE_SDAYNAME2& = &H2B
   LOCALE_SDAYNAME1& = &H2A
   LOCALE_SDATE& = &H1D
   LOCALE_SCURRENCY& = &H14
   LOCALE_SCOUNTRY& = &H6
   LOCALE_SENGCOUNTRY& = &H1002
   LOCALE_SENGLANGUAGE& = &H1001
   LOCALE_SABBREVMONTHNAME9& = &H4C
   LOCALE_SABBREVMONTHNAME8& = &H4B
   LOCALE_SABBREVMONTHNAME7& = &H4A
   LOCALE_SABBREVMONTHNAME6& = &H49
   LOCALE_SABBREVMONTHNAME5& = &H48
   LOCALE_SABBREVMONTHNAME4& = &H47
   LOCALE_SABBREVMONTHNAME3& = &H46
   LOCALE_SABBREVMONTHNAME2& = &H45
   LOCALE_SABBREVMONTHNAME13& = &H100F
   LOCALE_SABBREVMONTHNAME12& = &H4F
   LOCALE_SABBREVMONTHNAME11& = &H4E
   LOCALE_SABBREVMONTHNAME10& = &H4D
   LOCALE_SABBREVMONTHNAME1& = &H44
   LOCALE_SABBREVLANGNAME& = &H3
   LOCALE_SABBREVDAYNAME7& = &H37
   LOCALE_SABBREVDAYNAME6& = &H36
   LOCALE_SABBREVDAYNAME5& = &H35
   LOCALE_SABBREVDAYNAME4& = &H34
   LOCALE_SABBREVDAYNAME3& = &H33
   LOCALE_SABBREVDAYNAME2& = &H32
   LOCALE_SABBREVDAYNAME1& = &H31
   LOCALE_SABBREVCTRYNAME& = &H7
   LOCALE_ITLZERO& = &H25
   LOCALE_ITIME& = &H23
   LOCALE_IPOSSYMPRECEDES& = &H54
   LOCALE_IPOSSIGNPOSN& = &H52
   LOCALE_IPOSSEPBYSPACE& = &H55
   LOCALE_INEGSYMPRECEDES& = &H56
   LOCALE_INEGSIGNPOSN& = &H53
   LOCALE_INEGSEPBYSPACE& = &H57
   LOCALE_INEGCURR& = &H1C
   LOCALE_IMONLZERO& = &H27
   LOCALE_IMEASURE& = &HD
   LOCALE_ILDATE& = &H22
   LOCALE_ILANGUAGE& = &H1
   LOCALE_IINTLCURRDIGITS& = &H1A
   LOCALE_IDEFAULTLANGUAGE& = &H9
   LOCALE_IDEFAULTCOUNTRY& = &HA
   LOCALE_IDEFAULTCODEPAGE& = &HB
   LOCALE_IDAYLZERO& = &H26
   LOCALE_IDATE& = &H21
   LOCALE_ICURRENCY& = &H1B
   LOCALE_ICURRDIGITS& = &H19
   LOCALE_ICOUNTRY& = &H5
   LOCALE_ICENTURY& = &H24
   LOCALE_SYSTEM_DEFAULT& = &H800
   LOCALE_USER_DEFAULT& = &H400
   LOCALE_INEGNUMBER& = &H1010
End Enum

Public Enum LocaleFlags
   LCID_INSTALLED& = 1
   LCID_SUPPORTED& = 2
End Enum

Public Function FormatLocaleDateTime$(ByVal Locale As Long, _
                                      ByVal Expression As Date, _
                                      Optional ByVal NamedFormat As VbDateTimeFormat = vbGeneralDate)
    Dim myTime As SYSTEMTIME, s$, dl&, temp$

    If Locale = 0 Then Locale = LOCALE_SYSTEM_DEFAULT
    
    On Error GoTo ErrHandler
    
    ' --- Copy date/time info to SYSTEMTIME struct
    myTime.wMonth = Month(Expression)
    myTime.wDay = Day(Expression)
    myTime.wYear = Year(Expression)
    myTime.wDayOfWeek = Weekday(Expression)
    
    myTime.wHour = Hour(Expression)
    myTime.wMinute = Minute(Expression)
    myTime.wSecond = Second(Expression)
    myTime.wMilliseconds = 0&

    ' --- Clear the format buffer
    s$ = String$(255, Chr$(0))
    
    Select Case NamedFormat
    Case vbLongDate:
        dl& = GetDateFormat&(Locale, DATE_LONGDATE, myTime, vbNullString, s$, 254)
    Case vbShortDate:
        dl& = GetDateFormat&(Locale, DATE_SHORTDATE, myTime, vbNullString, s$, 254)
    Case vbLongTime:
        dl& = GetTimeFormat&(Locale, 0, myTime, 0, s$, 254)
    Case vbShortTime:
        dl& = GetTimeFormat&(Locale, TIME_NOSECONDS + TIME_NOTIMEMARKER + TIME_FORCE24HOURFORMAT, myTime, 0, s$, 254)
    Case Else: ' default: vbGeneralDate
        ' --- Assume no date if epoch date (30 DEC 1899)
        If Not (myTime.wMonth = 12 And myTime.wDay = 30 And myTime.wYear = 1899) Then
            dl& = GetDateFormat&(Locale, 0, myTime, vbNullString, s$, 254)
            If dl& > 0 Then temp$ = Left$(s$, dl& - 1) & " "
            s$ = String$(255, Chr$(0))
        End If
        ' --- Assume no time if midnight (fill in time anyway if temp$ is empty)
        If Not (myTime.wHour = 0 And myTime.wMinute = 0 And myTime.wSecond = 0) Or temp$ = "" Then
            dl& = GetTimeFormat&(Locale, 0, myTime, 0, s$, 254)
            If dl& > 0 Then
                s$ = temp$ & s$
                dl& = dl& + Len(temp$)
            End If
        Else
            s$ = temp$
        End If
    End Select

    If dl& = 0 Then GoTo ErrHandler
    FormatLocaleDateTime = Left$(s$, dl& - 1)
    Exit Function

ErrHandler:
    DoError (GetLastError())
End Function

Public Function FormatLocaleNumber$(ByVal Locale As Long, _
                                    ByVal Expression As Double, _
                                    Optional NumDigitsAfterDecimal As Long = -1, _
                                    Optional IncludeLeadingDigit As VbTriState = vbUseDefault, _
                                    Optional UseParensForNegativeNumbers As VbTriState = vbUseDefault, _
                                    Optional GroupDigits As VbTriState = vbUseDefault)
    Dim myFmt As NUMBERFMT, th$, dec$, mode&, sg$, ig&, dl&, n$

    If Locale = 0 Then Locale = LOCALE_SYSTEM_DEFAULT
    
    On Error GoTo ErrHandler

    '--- Defaults based on locale
    th$ = GetLocaleInfo(Locale, LOCALE_STHOUSAND)
    myFmt.lpThousandSep = th$
    dec$ = GetLocaleInfo(Locale, LOCALE_SDECIMAL)
    myFmt.lpDecimalSep = dec$
    
    '--- Number of digits after the decimal point
    If NumDigitsAfterDecimal >= 0 Then
        myFmt.NumDigits = NumDigitsAfterDecimal
    Else
        myFmt.NumDigits = CLng(GetLocaleInfo(Locale, LOCALE_IDIGITS))
    End If

    '--- Include leading 0 for fractions less than 1?
    If IncludeLeadingDigit <> vbUseDefault Then
        myFmt.LeadingZero = IIf(IncludeLeadingDigit = vbTrue, 1, 0)
    Else
        myFmt.LeadingZero = CLng(GetLocaleInfo(Locale, LOCALE_ILZERO))
    End If
    
    '--- Negative number mode
    mode& = CLng(GetLocaleInfo(Locale, LOCALE_INEGNUMBER))
    myFmt.NegativeOrder = IIf(UseParensForNegativeNumbers = vbTrue, 0, mode&)
    
    '--- Digit grouping (for thousands part only)
    sg$ = GetLocaleInfo(Locale, LOCALE_SGROUPING)
    ig& = CLng(Left$(sg$, InStr(sg$, ";") - 1))
    myFmt.Grouping = IIf(GroupDigits = vbFalse, 0, ig&)
    
    '--- Format it!
    n$ = String$(255, Chr$(0))
    dl& = GetNumberFormat(Locale, 0, CStr(Expression), myFmt, n$, 254)
    If dl& = 0 Then GoTo ErrHandler
    FormatLocaleNumber = Left$(n$, dl& - 1)
    Exit Function
    
ErrHandler:
    DoError (GetLastError())
End Function

Public Function FormatLocaleCurrency$(ByVal Locale As Long, _
                                      ByVal Expression As Double, _
                                      Optional NumDigitsAfterDecimal As Long = -1, _
                                      Optional IncludeLeadingDigit As VbTriState = vbUseDefault, _
                                      Optional UseParensForNegativeNumbers As VbTriState = vbUseDefault, _
                                      Optional GroupDigits As VbTriState = vbUseDefault)
    Dim myFmt As CURRENCYFMT, th$, dec$, cs$, mode&, sg$, ig&, dl&, n$

    If Locale = 0 Then Locale = LOCALE_SYSTEM_DEFAULT
    
    On Error GoTo ErrHandler

    '--- Defaults based on locale
    th$ = GetLocaleInfo(Locale, LOCALE_STHOUSAND)
    myFmt.lpThousandSep = th$
    dec$ = GetLocaleInfo(Locale, LOCALE_SMONDECIMALSEP)
    myFmt.lpDecimalSep = dec$
    cs$ = GetLocaleInfo(Locale, LOCALE_SCURRENCY)
    myFmt.lpCurrencySymbol = cs$
    myFmt.PositiveOrder = CLng(GetLocaleInfo(Locale, LOCALE_ICURRENCY))
    
    '--- Number of digits after the decimal point
    If NumDigitsAfterDecimal >= 0 Then
        myFmt.NumDigits = NumDigitsAfterDecimal
    Else
        myFmt.NumDigits = CLng(GetLocaleInfo(Locale, LOCALE_IINTLCURRDIGITS))
    End If

    '--- Include leading 0 for fractions less than 1?
    If IncludeLeadingDigit <> vbUseDefault Then
        myFmt.LeadingZero = IIf(IncludeLeadingDigit = vbTrue, 1, 0)
    Else
        myFmt.LeadingZero = CLng(GetLocaleInfo(Locale, LOCALE_ILZERO))
    End If
    
    '--- Negative number mode
    mode& = CLng(GetLocaleInfo(Locale, LOCALE_INEGCURR))
    myFmt.NegativeOrder = IIf(UseParensForNegativeNumbers = vbTrue, 0, mode&)
    
    '--- Digit grouping (for thousands part only)
    sg$ = GetLocaleInfo(Locale, LOCALE_SMONGROUPING)
    If InStr(sg$, ";") Then
       ig& = CLng(Left$(sg$, InStr(sg$, ";") - 1))
    Else
       ig& = CLng(sg$)
    End If
    myFmt.Grouping = IIf(GroupDigits = vbFalse, 0, ig&)
    
    '--- Format it!
    n$ = String$(255, Chr$(0))
    dl& = GetCurrencyFormat(Locale, 0, CStr(Expression), myFmt, n$, 254)
    If dl& = 0 Then GoTo ErrHandler
    FormatLocaleCurrency = Left$(n$, dl& - 1)
    Exit Function
    
ErrHandler:
    DoError (GetLastError())
End Function

'Public Function FormatMessage(ByVal Flags As Long, _
                              Source As Any, _
                              ByVal MessageId As Long, _
                              ByVal LanguageId As Long, _
                              ByVal Buffer As String, _
                              ByVal Size As Long, _
                              Arguments As Long)
'(ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long)
'End Function

Public Function GetLocaleInfo$(ByVal Locale As Long, _
                               ByVal LCType As Long)
    Dim s$, dl&
        
    If Locale = 0 Then Locale = LOCALE_SYSTEM_DEFAULT
    
    On Error GoTo ErrHandler:
    
    s$ = String$(255, Chr$(0))
    dl& = ntGetLocaleInfo(Locale, LCType, s$, 254)
    If dl& = 0 Then GoTo ErrHandler
    GetLocaleInfo = Left$(s$, dl& - 1)
    Exit Function

ErrHandler:
    DoError (GetLastError())
End Function

Public Function IsValidLocale(ByVal Locale As Long, _
                              Optional ByVal dwFlags As LocaleFlags = LCID_SUPPORTED) As Boolean
    IsValidLocale = (ntIsValidLocale(Locale, dwFlags) <> 0)
End Function

Public Function IsValidCodePage(ByVal CodePage As Long) As Boolean
    IsValidCodePage = (ntIsValidCodePage(CodePage) <> 0)
End Function

Public Function GetACP() As Long
    GetACP = ntGetACP()
End Function

Public Function ConvertDefaultLocale(ByVal Locale As Long)
    ConvertDefaultLocale = ntConvertDefaultLocale(Locale)
End Function

Public Function GetSystemDefaultLangID()
    GetSystemDefaultLangID = ntGetSystemDefaultLangID()
End Function

Public Function GetSystemDefaultLCID()
    GetSystemDefaultLCID = ntGetSystemDefaultLCID()
End Function

Public Function GetSystemTime() As Date
    Dim myTime As SYSTEMTIME, t As Date, d As Date
    ntGetSystemTime myTime
    t = TimeSerial(myTime.wHour, myTime.wMinute, myTime.wSecond)
    d = DateSerial(myTime.wYear, myTime.wMonth, myTime.wDay)
    GetSystemTime = t + d
End Function

'=== PRIVATE
Private Sub DoError(e&)
    Dim desc$
    Select Case e&
    Case ERROR_INSUFFICIENT_BUFFER:
        desc$ = "Insufficient buffer."
    Case ERROR_INVALID_FLAGS:
        desc$ = "Invalid flags."
    Case ERROR_INVALID_PARAMETER:
        desc$ = "Invalid parameter."
    Case Else:
        desc$ = "Unknown system error (" & Hex(e&) & ")."
    End Select
    Err.Raise vbObjectError, "NLS.Formats", desc$
End Sub

