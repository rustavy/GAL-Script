'Скрипт для импорта контактов в ПХ
'
'v 1.2.9 05.08.15 17:45 by ZRO@mail.ru
'
'=========================================================================================
'Option Explicit

CONST strServer                             = "tntex.logosgroup.pvt"
CONST strMailbox                            = "phgal"
CONST CdoPR_GIVEN_NAME                      = &H3A06001E  'First Name
CONST CdoPR_INITIALS                        = &H3A0A001E  'Initials
CONST CdoPR_SURNAME                         = &H3A11001E  'Last Name
CONST CdoPR_DISPLAY_NAME                    = &H3001001E  'Display Name
CONST CdoPR_ACCOUNT                         = &H3A00001E  'Alias
CONST CdoPR_TITLE                           = &H3A17001E  'Title
CONST CdoPR_COMPANY_NAME                    = &H3A16001E  'Company
CONST CdoPR_OFFICE_LOCATION                 = &H3A19001E  'Office
Const CdoPR_COMPANY_MAIN_PHONE_NUMBER       = &H3A57001E
Const CdoPR_PRIMARY_TELEPHONE_NUMBER        = &H3A1A001E
CONST CdoPR_HOME_TELEPHONE_NUMBER           = &H3A09001E  'Phone
CONST CdoPR_HOME2_TELEPHONE_NUMBER          = &H3A2F001E  'Home Phone 2
CONST CdoPR_HOME_FAX_NUMBER                 = &H3A25001E  'Home Fax
CONST CdoPR_HOME_ADDRESS_STREET             = &H3A5D001E  'Address
CONST CdoPR_HOME_ADDRESS_CITY               = &H3A59001E  'City
CONST CdoPR_HOME_ADDRESS_STATE_OR_PROVINCE  = &H3A5C001E  'State
CONST CdoPR_HOME_ADDRESS_POSTAL_CODE        = &H3A5B001E  'Zip
CONST CdoPR_HOME_ADDRESS_COUNTRY            = &H3A5A001E  'Country
CONST CdoPR_MANAGER_NAME                    = &H3A4E001E  'Manager
CONST CdoPR_OFFICE_TELEPHONE_NUMBER         = &H3A08001E  'Business
                                                          'Phone
CONST CdoPR_OFFICE2_TELEPHONE_NUMBER        = &H3A1B001E  'Business
                                                          'Phone 2
CONST CdoPR_BUSINESS_FAX_NUMBER             = &H3A24001E  'Fax
CONST CdoPR_ASSISTANT                       = &H3A30001E  'Assistant
CONST CdoPR_ASSISTANT_TELEPHONE_NUMBER      = &H3A2E001E  'Asistant
                                                          'Phone Number
CONST CdoPR_MOBILE_TELEPHONE_NUMBER         = &H3A1C001E  'Mobile
CONST CdoPR_PAGER_TELEPHONE_NUMBER          = &H3A21001E  'Pager
Const CdoPR_DEPARTMENT_NAME                 = &H3A18001E  'Departament

Const CdoPR_CUSTOM_ATTRIBUTE_1              = &H802D001E  'Custom Attribute 1
Const CdoPR_CUSTOM_ATTRIBUTE_2              = &H802E001E  'Custom Attribute 2
Const CdoPR_CUSTOM_ATTRIBUTE_3              = &H802F001E  'Custom Attribute 3
Const CdoPR_CUSTOM_ATTRIBUTE_9              = &H8035001E  'Custom Attribute 9
Const CdoPR_HIDE_FROM_ADDRESS_BOOK          = &H80B9000B  'True = Display
                                                          'False = Hide
Const PR_EMS_AB_PROXY_ADDRESSES             = &H800F101E  'E-mail Addresses
Const PR_SMTP_ADDRESS                       = &H39FE001E  'Primary SMTP Address
Const PR_EMS_AB_ORGANIZATIONAL_UNIT_NAME    = &H8102101E
Const PR_EMS_AB_TELEPHONE_NUMBER            = &H8012101E
Const PR_OTHER_TELEPHONE_NUMBER             = &H3A1F001E
Const PR_BUSINESS2_TELEPHONE_NUMBER_A_MV    = &H3A1B101E

Const ADS_PROPERTY_CLEAR   = 1
Const ADS_PROPERTY_UPDATE  = 2
Const ADS_PROPERTY_APPEND  = 3
Const ADS_PROPERTY_DELETE  = 4

'Const logg As Integer = 3 ' 1 - только в файл, 2 - только на экран, 3 - и в файл и на экран
Const logg  = 3 ' 1 - только в файл, 2 - только на экран, 3 - и в файл и на экран

Const strObject = "contact"

Dim objSession
Dim objAddrEntries
Dim objAddressEntry
Dim objField
Dim objFilter
Dim fso
Dim objOu
Dim objUser
Dim aclsContactList
Dim sAMAccountName
Dim objRootLDAP
Dim objContainer
Dim objContact

Dim strProfileInfo, strTimeStamp
Dim iCount, eCount, xCount
Dim sCurPath, strArg
Dim strFirstName, strLastName, strAlias, strEmail, strDisName, strMainDefault, strProxy, strMBName, strMailbox1
Dim strDepartment, strCompany, strDescription, strPhone, strOffice, strlExDN, strTDisName
Dim strCustomAttribute1, strCustomAttribute2, strCustomAttribute3, strCustomAttribute9
Dim strCcn, strCMail, strCgivenName, strCsn, strCDisName, strCeAttr, strCPhone, strCOffice, strClExDN

On Error Resume Next
strArg = Trim(UCase(WScript.Arguments.Item(0)))
On Error Goto 0
strProfileInfo = strServer & vbLf & strMailbox

'-------------------------------------------------------------'
'Получаем список адресов из Глобальной Адресной Книги Exchange
'-------------------------------------------------------------'
Set objSession = CreateObject("MAPI.Session")
objSession.Logon , , False, False, , True, strProfileInfo 'Логинимся к серверу...
Set objAddrEntries = objSession.AddressLists("Global Address List").AddressEntries  'Global Address List  Глобальный список адресов

'-------------------------------------------------------------'
'Создаем файл лога если указано писать в файл
'-------------------------------------------------------------'
if logg=1 or logg=3 then
  Set fso = CreateObject("Scripting.FileSystemObject")
  sCurPath = fso.GetParentFolderName(Wscript.ScriptFullName) & "\logos2ph.log"

  On Error Resume Next
  fso.deleteFile sCurPath ' Удаляем старый лог файл.
  On Error Goto 0
end if

'-------------------------------------------------------------'
'Функция логирования
'-------------------------------------------------------------'
Sub log(sData)
  Dim ts, ForAppending
  strTimeStamp = CStr(Now())
'  strTimeStamp = Replace(strTimeStamp, "/", "")
'  strTimeStamp = Replace(strTimeStamp, " ", "")
'  strTimeStamp = Replace(strTimeStamp, ":", "")
  if logg=1 or logg=3 then
    ForAppending = 8
    Set ts = fso.OpenTextFile(sCurPath, ForAppending, True)
    ts.Write strTimeStamp & " " & sData & chr(13) & chr(10)
    ts.Close
  end if
  if logg >1 then
    if strArg <> "/Q" then
      wscript.echo strTimeStamp & " " & sData
    end if
  end if
  strTimeStamp = ""
End Sub

'-------------------------------------------------------------'
'Important create OU=Contacts or change value for strContainer
'-------------------------------------------------------------'
set objOU =GetObject("LDAP://ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt")
Set aclsContactList = CreateObject("Scripting.Dictionary")

'!!!!!ВНИМАНИЕ. Так делаем сейчас, т.е. добавляются только новые контакты.
' Собираем список имеющихся контактов
strMBName = ""
Log ("Собираем список имеющихся контактов...")
For Each objUser In objOU
  If objUser.class = strObject Then            ' Каждый обьект класса "contact" добавляется в список
    strMBName = Trim(mid(objUser.legacyExchangeDN, instrrev(objUser.legacyExchangeDN,"=")+1))
    aclsContactList.Add objUser.cn, strMBName  ' Список выглядит как "cn" "strMBName"
    Log(objUser.cn & " Хранится как: " & aclsContactList.Item(objUser.cn))
  End If
  strMBName = ""
Next
Set objOU = Nothing
Log ("Найдено   : " & aclsContactList.Count)
Log ("")

'-------------------------------------------------------------'
'Задаем OU в которой будут создаваться контакты
'-------------------------------------------------------------'
Set objRootLDAP = GetObject("LDAP://rootDSE")

On Error Goto 0
Set objContainer = GetObject("LDAP://ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt")
On Error Resume Next

'-------------------------------------------------------------'
'Сбрасывам переменные
'-------------------------------------------------------------'
iCount               = 0
eCount               = 0
xCount               = 0
strCcn               = ""
strCMail             = ""
strCsn               = ""
strCDisName          = ""
strCeAttr            = ""
strmAPIRecipient     = false
strinternetEncoding  = 1310720
strCmAPIRecipient    = ""
strCinternetEncoding = ""
strTDisName          = ""

'-------------------------------------------------------------'
'Начинаем проверять
'-------------------------------------------------------------'
log("Начинаем проверять...")

'-------------------------------------------------------------'
'Для каждого контакта в адресной книге выполняем следующее
'-------------------------------------------------------------'
For Each objAddressEntry In objAddrEntries
'-------------------------------------------------------------'
'Если в поле legacyExchangeDN содержится /O=LOGOS-M/OU=LGROUP/cn=Recipients/cn=" & strMBName
'                                                      ^^^^^^
'-------------------------------------------------------------'
  if mid(objAddressEntry.Address,15,6)="LGROUP" then
    On Error Resume Next

'-------------------------------------------------------------'
'Извлекаем данные контакта из адресной книги
'-------------------------------------------------------------'
    strFirstName   = Trim(objAddressEntry.Fields(CdoPR_GIVEN_NAME).Value)
    strLastName    = Trim(objAddressEntry.Fields(CdoPR_SURNAME).Value)
'    strFirstName   = Trim(CStr(objAddressEntry.Fields(CdoPR_GIVEN_NAME).Value))
'    strLastName    = Trim(CStr(objAddressEntry.Fields(CdoPR_SURNAME).Value))
    strAlias       = Trim(objAddressEntry.Fields(CdoPR_ACCOUNT).Value)

'-------------------------------------------------------------'
'Выясняем основной SMTP адрес
'-------------------------------------------------------------'
    Set objField = objAddressEntry.Fields(PR_EMS_AB_PROXY_ADDRESSES)

    For Each v In objField.Value
      If Mid(v, 1, 4) = "SMTP" Then
        strEmail = Trim(Mid(v, 6))
        strMainDefault = Trim(v)
        strProxy = Trim(v)
      End If
    Next

'-------------------------------------------------------------'
'Запрашиваем остальные аттрибуты
'-------------------------------------------------------------'
    strMBName           = Trim(mid(objAddressEntry.Address, instrrev(objAddressEntry.Address,"=")+1)) 'Alias
    strlExDN            = Trim(CStr(objAddressEntry.Address))                                         'legacyExchangeDN натуральный
    strMailbox1         = Trim("/O=LOGOS-M/OU=LGROUP/cn=Recipients/cn=" & strMBName)                  'legacyExchangeDN самодельный
    strDisName          = Trim(objAddressEntry.Fields(CdoPR_DISPLAY_NAME).Value)                      'Отображаемое имя
    strDepartment       = Trim(objAddressEntry.Fields(CdoPR_DEPARTMENT_NAME).Value)                   'Департамент
    strCompany          = Trim(objAddressEntry.Fields(CdoPR_COMPANY_NAME).Value)                      'Компания
    strDescription      = Trim(objAddressEntry.Fields(CdoPR_TITLE).Value)                             'Должность
    strPhone            = Trim(CStr(objAddressEntry.Fields(CdoPR_OFFICE_TELEPHONE_NUMBER).Value))     'Номер телефона
    strOffice           = Trim(CStr(objAddressEntry.Fields(CdoPR_OFFICE_LOCATION).Value))             'Офис
    strCustomAttribute1 = Trim(objAddressEntry.Fields(CdoPR_CUSTOM_ATTRIBUTE_1).Value)                'Аттрибут 1
    strCustomAttribute2 = Trim(objAddressEntry.Fields(CdoPR_CUSTOM_ATTRIBUTE_2).Value)                'Аттрибут 2
    strCustomAttribute3 = Trim(objAddressEntry.Fields(CdoPR_CUSTOM_ATTRIBUTE_3).Value)                'Аттрибут 3
    strCustomAttribute9 = strMBName                                                                   'Аттрибут 9

    strTDisName = strDisName
    strTDisName = replace(strTDisName,"\","\\")
    strTDisName = replace(strTDisName,"/","\/")
    strTDisName = replace(strTDisName,"#","\#")
    strTDisName = replace(strTDisName,",","\,")

'    On Error Goto 0

'-------------------------------------------------------------'
'Сверяем
'-------------------------------------------------------------'
    If strEmail <> "stp@logosgroup.ru" and _
       strEmail <> "stp@finans-media.ru" and _
       strEmail <> "epak@mail.ru" Then ' Пропускаем исключения.
      If (aclsContactList.Exists(strMBName) or aclsContactList.Exists(strDisName)) and _
         Not (aclsContactList.Exists(strMBName) and aclsContactList.Exists(strDisName)) Then ' Если контакт есть в списке - пропускаем.
'        Set objContact = GetObject("LDAP://" & "cn=" & strMBName & ",OU=Press-House-AL,OU=Logos,DC=logosgroup,DC=pvt")
        If aclsContactList.Exists(strMBName) then
          Err.Clear
          Set objContact = GetObject("LDAP://" & "cn=" & strMBName & ",ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt")
          If (Err.number <> 0) Then
            log("Попытались открыть: LDAP://" & "cn=" & strMBName & ",ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt - Что-то пошло не так...")
            log("Ошибка: " & Err.Number)
            Select Case Err.Number
            Case 424
              log("Не смогли создать обьект.")
            Case &h80072032
              log("LDAP_INVALID_DN_SYNTAX This error occurs when a distinguished name used for the creation of objects contains invalid characters.")
            Case &h80071392
              log("Object with this name already exists.")
            Case &h80072030
              log("LDAP_NO_SUCH_OBJECT This error is similar to ADS_BAD_PATHNAME (0x80005008) - during the BIND process, an LDAP object path was passed from a non existing object.")
            Case &h80072030
              log("Нет такого контакта")
              Wscript.Echo "OU doesn't exist"
            Case Else
              log(Err.Descritption)
            End Select
'            Wscript.Quit 2
          End If
        end if
        If aclsContactList.Exists(strDisName) then
          Err.Clear
          Set objContact = GetObject("LDAP://" & "cn=" & strTDisName & ",ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt")
          If (Err.number <> 0) Then
            log("Попытались открыть: LDAP://" & "cn=" & strTDisName & ",ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt - Что-то пошло не так...")
            log("Ошибка: " & Err.Number)
            Select Case Err.Number
            Case 424
              log("Не смогли создать обьект.")
            Case &h80072032
              log("LDAP_INVALID_DN_SYNTAX This error occurs when a distinguished name used for the creation of objects contains invalid characters.")
            Case &h80071392
              log("Object with this name already exists.")
            Case &h80072030
              log("LDAP_NO_SUCH_OBJECT This error is similar to ADS_BAD_PATHNAME (0x80005008) - during the BIND process, an LDAP object path was passed from a non existing object.")
            Case &h80072030
              log("Нет такого контакта")
              Wscript.Echo "OU doesn't exist"
            Case Else
              log(Err.Descritption)
            End Select
'            Wscript.Quit 2
          End If
        end if

        Err.Clear
        objContact.GetInfo
        If (Err.number <> 0) Then
          log("Попытались запросить данные контакта - Что-то пошло не так...")
          log("Ошибка: " & Err.Number)
          Select Case Err.Number
          Case 424
            log("Не смогли создать обьект.")
          Case &h80072032
            log("LDAP_INVALID_DN_SYNTAX This error occurs when a distinguished name used for the creation of objects contains invalid characters.")
          Case &h80071392
            log("Object with this name already exists.")
          Case &h80072030
            log("LDAP_NO_SUCH_OBJECT This error is similar to ADS_BAD_PATHNAME (0x80005008) - during the BIND process, an LDAP object path was passed from a non existing object.")
          Case &h80072030
            log("Нет такого контакта")
            Wscript.Echo "OU doesn't exist"
          Case Else
            log(Err.Descritption)
          End Select
'          Wscript.Quit 2
        End If

        strCcn               = Trim(LCase(CStr(objContact.Get("cn"))))
        strCMail             = Trim(LCase(CStr(objContact.Get("Mail"))))
        strClExDN            = Trim(LCase(CStr(objContact.Get("legacyExchangeDN"))))
        strCgivenName        = Trim(LCase(CStr(objContact.Get("givenName"))))
        strCsn               = Trim(LCase(CStr(objContact.Get("sn"))))
        strCDisName          = Trim(LCase(CStr(objContact.Get("displayName"))))
        strCeAttr            = Trim(LCase(CStr(objContact.Get("extensionAttribute1"))))
        strCPhone            = Trim(LCase(CStr(objContact.Get("telephoneNumber"))))
        strCOffice           = Trim(LCase(CStr(objContact.Get("physicalDeliveryOfficeName"))))
        strCmAPIRecipient    = Trim(LCase(CStr(objContact.Get("mAPIRecipient"))))
        strCinternetEncoding = Trim(LCase(CStr(objContact.Get("internetEncoding"))))

        if LCase(strMBName)<>strCcn Then
          log("Искали: `" & strMBName & "` Нашли: `" & strCcn & "`")
        end if

        If strCMail<>LCase(strEmail) or _
           strClExDN<>LCase(strlExDN) or _
           (strCgivenName<>LCase(strFirstName) and strCgivenName<>"") or _
           (strCsn<>LCase(strLastName) and strCsn<>"") or _
           (strCDisName<>LCase(strDisName) and strCDisName<>"") or _
           (LCase(strDisName)<>strCcn and strCcn<>"") or _
           (strCeAttr<>LCase(strCustomAttribute1) and strCeAttr<>"нпдп") or _
           strCPhone<>LCase(strPhone) or _
           strCmAPIRecipient<>LCase(strmAPIRecipient) or _
           strCinternetEncoding<>LCase(strinternetEncoding) or _
           strCOffice<>LCase(strOffice) Then

          log("")
          log("Контакту нужны обновления: " & strEmail)
          If strDisName<>strCcn and strCcn<>"" then
            log("Контакту необходимо обновить cn: `" & strCcn & "`>`" & strDisName & "`")

            Set objPContact = GetObject("LDAP://ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt")

            Err.Clear
            Set objNContact = objPContact.MoveHere("LDAP://" & "cn=" & strMBName & ",ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt", UnEscape("cn=" & strTDisName))
            If (Err.number <> 0) Then
              log("Что-то пошло не так...")
              log("Ошибка: " & Err.Number)
              Select Case Err.Number
              Case 424
                log("Не смогли создать обьект. " & "cn=" & strTDisName)
              Case &h80072032
                log("LDAP_INVALID_DN_SYNTAX This error occurs when a distinguished name used for the creation of objects contains invalid characters.")
              Case &h80071392
                log("Object with this name already exists.")
              Case &h80072030
                log("LDAP_NO_SUCH_OBJECT This error is similar to ADS_BAD_PATHNAME (0x80005008) - during the BIND process, an LDAP object path was passed from a non existing object.")
              Case &h80072030
                log("Нет такого контакта")
                Wscript.Echo "OU doesn't exist"
              Case Else
                log(Err.Descritption)
              End Select
'              Wscript.Quit 2
            End If

            Err.Clear
'   On Error Goto 0
            objNContact.SetInfo

            If (Err.number <> 0) Then
              log("Не можем ничего поделать, контакт с таким именем уже есть в конечной OU.")
              log("Ошибка: " & Err.Number)
              Select Case Err.Number
              Case &h80072032
                log("LDAP_INVALID_DN_SYNTAX This error occurs when a distinguished name used for the creation of objects contains invalid characters.")
              Case &h80071392
                log("Object with this name already exists.")
              Case &h80072030
                log("LDAP_NO_SUCH_OBJECT This error is similar to ADS_BAD_PATHNAME (0x80005008) - during the BIND process, an LDAP object path was passed from a non existing object.")
              Case &h80072030
                log("Нет такого контакта")
                Wscript.Echo "OU doesn't exist"
              Case Else
                log(Err.Descritption)
              End Select
'              Wscript.Quit 2
            End If
'   On Error Resume Next

            Set objNContact = Nothing
            Set objPContact = Nothing

            Set objContact  = Nothing
            Set objContact  = GetObject("LDAP://" & "cn=" & strTDisName & ",ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt")
            objContact.GetInfo

            strCcn               = Trim(LCase(CStr(objContact.Get("cn"))))
            strCMail             = Trim(LCase(CStr(objContact.Get("Mail"))))
            strClExDN            = Trim(LCase(CStr(objContact.Get("legacyExchangeDN"))))
            strCgivenName        = Trim(LCase(CStr(objContact.Get("givenName"))))
            strCsn               = Trim(LCase(CStr(objContact.Get("sn"))))
            strCDisName          = Trim(LCase(CStr(objContact.Get("displayName"))))
            strCeAttr            = Trim(LCase(CStr(objContact.Get("extensionAttribute1"))))
            strCPhone            = Trim(LCase(CStr(objContact.Get("telephoneNumber"))))
            strCOffice           = Trim(LCase(CStr(objContact.Get("physicalDeliveryOfficeName"))))
            strCmAPIRecipient    = Trim(LCase(CStr(objContact.Get("mAPIRecipient"))))
            strCinternetEncoding = Trim(LCase(CStr(objContact.Get("internetEncoding"))))

            if LCase(strDisName)<>strCcn Then
              log("Искали: `" & strDisName & "` Нашли: `" & strCcn & "`")
            end if

          end if
          If strCMail<>LCase(strEmail) then
            log("Контакту необходимо обновить адрес: `" & strCMail & "`>`" & LCase(strEmail) & "`")
          end if
          If strCgivenName<>LCase(strFirstName) and strCgivenName<>"" then
            log("Контакту необходимо обновить имя: `" & strCgivenName & "`>`" & LCase(strFirstName) & "`")
          end if
          If strCsn<>LCase(strLastName) and strCsn<>"" then
            log("Контакту необходимо обновить фамилию: `" & strCsn & "`>`" & LCase(strLastName) & "`")
          end if
          If strCDisName<>LCase(strDisName) then
            log("Контакту необходимо обновить отображение: `" & strCDisName & "`>`" & LCase(strDisName) & "`")
          end if
          If strCeAttr<>LCase(strCustomAttribute1) and strCeAttr<>"нпдп" then
            log("Контакту необходимо обновить аттрибут: `" & strCeAttr & "`>`" & LCase(strCustomAttribute1) & "`")
          end if
          If strCPhone<>LCase(strPhone) then
            log("Контакту необходимо обновить тлф   : `" & strCPhone & "`>`" & LCase(strPhone) & "`")
          end if
          If strCOffice<>LCase(strOffice) then
            log("Контакту необходимо обновить офис  : `" & strCOffice & "`>`" & LCase(strOffice) & "`")
          end if

          if strDisName<>strCcn then
            log("cn               : " & strDisName)
          end if
          if strFirstName <> "" then
            objContact.Put "givenName", UnEscape(strFirstName)
            log("givenName        : " & strFirstName)
          else
            objContact.PutEx ADS_PROPERTY_CLEAR, "givenName", vbNullString
            log("givenName        : " & strFirstName)
          end if
          if strLastName <> "" then
            objContact.Put "sn", UnEscape(strLastName)
            log("sn               : " & strLastName)
          else
            objContact.PutEx ADS_PROPERTY_CLEAR, "sn", vbNullString
            log("sn               : " & strLastName)
          end if
          if strEmail <> "" then
            objContact.Put "Mail", UnEscape(strEmail)
            log("Mail             : " & strEmail)
          end if
          if strProxy <> "" then
            objContact.Put "proxyAddresses", UnEscape(strProxy)
            log("proxyAddresses   : " & strProxy)
          end if
          if strMainDefault <> "" then
            objContact.Put "targetAddress", UnEscape(strMainDefault)
            log("targetAddress    : " & strMainDefault)
          end if
          if strlExDN <> "" then
            objContact.Put "legacyExchangeDN", UnEscape(strlExDN)
            log("legacyExchangeDN : " & strClExDN)
            log("legacyExchangeDN : " & strlExDN)
          end if
'-------------------
          if strCmAPIRecipient<>LCase(strmAPIRecipient) then
            log("Контакту необходимо обновить mAPIRecipient : `" & strCmAPIRecipient & "`>`" & LCase(strmAPIRecipient) & "`")
            objContact.Put "mAPIRecipient", strmAPIRecipient
            log("mAPIRecipient    : " & strmAPIRecipient)
          end if
          if strCinternetEncoding<>LCase(strinternetEncoding) then
            log("Контакту необходимо обновить internetEncoding : `" & strCinternetEncoding & "`>`" & LCase(strinternetEncoding) & "`")
            objContact.Put "internetEncoding", strinternetEncoding
            log("internetEncoding : " & strinternetEncoding)
          end if
'-------------------
          if strAlias <> "" then
            objContact.Put "mailNickname", UnEscape(strAlias)
            log("mailNickname     : " & strAlias)
          end if
          if strDisName <> "" then
            objContact.Put "displayName", UnEscape(strDisName)
            log("displayName      : " & strDisName)
          else
            objContact.PutEx ADS_PROPERTY_CLEAR, "displayName", vbNullString
            log("displayName      : " & strDisName)
          end if
          if strDepartment <> "" Then
            objContact.Put "department", UnEscape(strDepartment)
            log("department       : " & strDepartment)
          else
            objContact.PutEx ADS_PROPERTY_CLEAR, "department", vbNullString
            log("department       : " & strDepartment)
          end if
          If strCompany <> "" Then
            objContact.Put "company", UnEscape(strCompany)
            log("company          : " & strCompany)
          else
            objContact.Put "company", "Группа компаний 'Логос'"
            log("company          : Группа компаний 'Логос'")
          end if
          If strDescription <> "" Then
            objContact.Put "description", UnEscape(strDescription)
            objContact.Put "title", UnEscape(strDescription)
            log("description      : " & strDescription)
          else
            objContact.Put "description", "Группа компаний 'Логос'"
            objContact.Put "title", "Группа компаний 'Логос'"
            log("description      : Группа компаний 'Логос'")
          end if
          if strPhone <> "" Then
            if strCPhone <> "" Then
              objContact.PutEx ADS_PROPERTY_UPDATE, "telephoneNumber", Array(strPhone)
            else
              objContact.Put "telephoneNumber", Array(strPhone)
            end if
            log("telephoneNumber  : " & strPhone)
          else
            objContact.PutEx ADS_PROPERTY_CLEAR, "telephoneNumber", vbNullString
            log("telephoneNumber  : " & strPhone)
          end if
          if strOffice <> "" Then
            if strCOffice <> "" Then
              objContact.Put "physicalDeliveryOfficeName", UnEscape(strOffice)
            else
              objContact.Put "physicalDeliveryOfficeName", strOffice
            end if
            log("physicalDeliveryOfficeName: " & strOffice)
          else
            objContact.PutEx ADS_PROPERTY_CLEAR, "physicalDeliveryOfficeName", vbNullString
            log("physicalDeliveryOfficeName: " & strOffice)
          end if

          if strCustomAttribute1 <> "" then
            if strCeAttr <> "" then
              objContact.Put "extensionAttribute1", UnEscape(strCustomAttribute1)
            else
              objContact.Put "extensionAttribute1", UnEscape(strCustomAttribute1)
            end if
            log("Custom Attribute1: " & strCustomAttribute1)
          else
            objContact.Put "extensionAttribute1", "НПДП"
            log("Custom Attribute1: НПДП")
          end if
          if strCustomAttribute2 <> "" then
            objContact.Put "extensionAttribute2", UnEscape(strCustomAttribute2)
            log("Custom Attribute2: " & strCustomAttribute2)
          end if
          if strCustomAttribute3 <> "" then
            objContact.Put "extensionAttribute3", UnEscape(strCustomAttribute3)
            log("Custom Attribute3: " & strCustomAttribute3)
          end if
'          ' Записываем в 9й атрибут ID контакта
'          if strCustomAttribute9 <> "" then
'            objContact.Put "extensionAttribute9", UnEscape(strCustomAttribute9)
'            log("Custom Attribute9: " & strCustomAttribute9)
'          end if

          Err.Clear
'   On Error Goto 0
          objContact.SetInfo

          If (Err.number <> 0) Then
            log("Попытались записать данные - Что-то пошло не так...")
            log("Ошибка: " & Err.Number)
            Select Case Err.Number
            Case 424
              log("Не смогли создать обьект.")
            Case &h80072032
              log("LDAP_INVALID_DN_SYNTAX This error occurs when a distinguished name used for the creation of objects contains invalid characters.")
            Case &h80071392
              log("Object with this name already exists.")
            Case &h80072030
              log("LDAP_NO_SUCH_OBJECT This error is similar to ADS_BAD_PATHNAME (0x80005008) - during the BIND process, an LDAP object path was passed from a non existing object.")
            Case &h80072030
              log("Нет такого контакта")
'              Wscript.Echo "OU doesn't exist"
            Case Else
              log(Err.Descritption)
            End Select
'            Wscript.Quit 2
          End If
'   On Error Resume Next
'wscript.quit

          log("")
          xCount = xCount + 1
        else
          log("Контакт существует, пропущен: " & strMBName & " (" & strDisName & ")")
          eCount = eCount + 1
        end if
        aclsContactList.Remove(strMBName)
        aclsContactList.Remove(strDisName)
'------------------------------------------------------------------------------
      ElseIf Not aclsContactList.Exists(strMBName) and Not aclsContactList.Exists(strDisName) then ' Если контакта нет в списке - создаем.
        if instrrev(strEmail,"@logosgroup.ru") <> 0 or _
           instrrev(strEmail,"@read.ru") <> 0 or _
           instrrev(strEmail,"@centropechat.ru") <> 0 or _
           instrrev(strEmail,"@presslogist.ru") <> 0 or _
           instrrev(strEmail,"@mediadis.ru") <> 0 or _
           instrrev(strEmail,"@inkorinvest.ru") <> 0 or _
           instrrev(strEmail,"@logos-finance.ru") <> 0 or _
           instrrev(strEmail,"@kpechati.ru") <> 0 or _
           instrrev(strEmail,"@finans-media.ru") <> 0 or _
           instrrev(strEmail,"@center-srv.ru") <> 0 or _
           instrrev(strEmail,"@idlogos.com") <> 0 or _
           instrrev(strEmail,"@idlogos.ru") <> 0 or _
           instrrev(strEmail,"@invest-base.ru") <> 0 or _
           instrrev(strEmail,"@npdpgroup.ru") <> 0 or _
           instrrev(strEmail,"@ff.ff") <> 0 or _
           instrrev(strEmail,"@konsult-center.ru") <> 0 then ' Если у контакта адрес из LOGOSGROUP
          log("")
          log("Добавляем контакт: " & strMBName & " (" & strDisName & ")")
'          log("Добавляем контакт: " & strLastName & " " & strFirstName)

          Err.Clear
   On Error Goto 0
'          Set objContact = objContainer.Create("Contact",UnEscape("cn=" & strTDisName))
          Set objContainer = Nothing
          Set objContainer = GetObject("LDAP://ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt")
          Set objContact = Nothing
          Set objContact = objContainer.Create("Contact", "cn=" & strTDisName)
   On Error Resume Next
'          Set objContact = objContainer.Create("Contact","cn=" & strLastName & " " & strFirstName)

          If (Err.number <> 0) Then
            log("Попытались создать контакт - Что-то пошло не так...")
            log("Ошибка: " & Err.Number)
            Select Case Err.Number
            Case &h80072032
              log("LDAP_INVALID_DN_SYNTAX This error occurs when a distinguished name used for the creation of objects contains invalid characters.")
            Case &h80071392
              log("Object with this name already exists.")
            Case &h80072030
              log("LDAP_NO_SUCH_OBJECT This error is similar to ADS_BAD_PATHNAME (0x80005008) - during the BIND process, an LDAP object path was passed from a non existing object.")
            Case &h80072030
              log("Нет такого контакта")
              Wscript.Echo "OU doesn't exist"
            Case 424
              log("Не смогли создать обьект контакта. " & "cn=" & strTDisName)
            Case Else
              log(Err.Descritption)
            End Select
'            Wscript.Quit 2
          End If

          if strFirstName <> "" then
            objContact.Put "givenName", UnEscape(strFirstName)
            log("givenName        : " & strFirstName)
          end if
          if strLastName <> "" then
            objContact.Put "sn", UnEscape(strLastName)
            log("sn               : " & strLastName)
          end if
          if strEmail <> "" then
            objContact.Put "Mail", UnEscape(strEmail)
            log("Mail             : " & strEmail)
          end if
          if strProxy <> "" then
            objContact.Put "proxyAddresses", UnEscape(strProxy)
            log("proxyAddresses   : " & strProxy)
          end if
          if strMainDefault <> "" then
            objContact.Put "targetAddress", UnEscape(strMainDefault)
            log("targetAddress    : " & strMainDefault)
          end if
          if strlExDN <> "" then
            objContact.Put "legacyExchangeDN", UnEscape(strlExDN)
            log("legacyExchangeDN : " & strlExDN)
          end if
'-------------------
'          if strCmAPIRecipient<>LCase(strmAPIRecipient) then
            objContact.Put "mAPIRecipient", strmAPIRecipient
            log("mAPIRecipient    : " & strmAPIRecipient)
'          end if
'          if strCinternetEncoding<>LCase(strinternetEncoding) then
            objContact.Put "internetEncoding", strinternetEncoding
            log("internetEncoding : " & strinternetEncoding)
'          end if
'-------------------
          if strAlias <> "" then
            objContact.Put "mailNickname", UnEscape(strAlias)
            log("mailNickname     : " & strAlias)
          end if
          if strDisName <> "" then
            objContact.Put "displayName", UnEscape(strDisName)
            log("displayName      : " & strDisName)
          end if
          if strDepartment <> "" Then
            objContact.Put "department", UnEscape(strDepartment)
            log("department       : " & strDepartment)
          end if
          If strCompany <> "" Then
            objContact.Put "company", UnEscape(strCompany)
            log("company          : " & strCompany)
          else
            objContact.Put "company", "Группа компаний 'Логос'"
            log("company          : Группа компаний 'Логос'")
          end if
          If strDescription <> "" Then
            objContact.Put "description", UnEscape(strDescription)
            objContact.Put "title", UnEscape(strDescription)
            log("description      : " & strDescription)
          else
            objContact.Put "description", "Группа компаний 'Логос'"
            objContact.Put "title", "Группа компаний 'Логос'"
            log("description      : Группа компаний 'Логос'")
          end if
          if strPhone <> "" Then
            objContact.Put "telephoneNumber", UnEscape(strPhone)
            log("telephoneNumber  : " & strPhone)
          end if
          if strOffice <> "" Then
            objContact.Put "physicalDeliveryOfficeName", UnEscape(strOffice)
            log("physicalDeliveryOfficeName: " & strOffice)
          end if

          if strCustomAttribute1 <> "" then
            objContact.Put "extensionAttribute1", UnEscape(strCustomAttribute1)
            log("Custom Attribute1: " & strCustomAttribute1)
          end if
          if strCustomAttribute2 <> "" then
            objContact.Put "extensionAttribute2", UnEscape(strCustomAttribute2)
            log("Custom Attribute2: " & strCustomAttribute2)
          end if
          if strCustomAttribute3 <> "" then
            objContact.Put "extensionAttribute3", UnEscape(strCustomAttribute3)
            log("Custom Attribute3: " & strCustomAttribute3)
          end if
'          ' Записываем в 9й атрибут ID контакта
'          if strCustomAttribute9 <> "" then
'            objContact.Put "extensionAttribute9", UnEscape(strCustomAttribute9)
'            log("Custom Attribute9: " & strCustomAttribute9)
'          end if

          Err.Clear
          objContact.SetInfo
          If (Err.number <> 0) Then
            log("Не смогли добавить контакт.")
            log("Ошибка: " & Err.Number)
            Select Case Err.Number
            Case 424
              log("Не смогли создать обьект.")
            Case &h80072032
              log("LDAP_INVALID_DN_SYNTAX This error occurs when a distinguished name used for the creation of objects contains invalid characters.")
            Case &h80071392
              log("Object with this name already exists.")
            Case &h80072030
              log("LDAP_NO_SUCH_OBJECT This error is similar to ADS_BAD_PATHNAME (0x80005008) - during the BIND process, an LDAP object path was passed from a non existing object.")
            Case &h80072030
              log("Нет такого контакта")
              Wscript.Echo "OU doesn't exist"
            Case Else
              log(Err.Descritption)
            End Select
'            Wscript.Quit 2
          End If

          log("")

          iCount = iCount + 1
        else
          log("У пользователя внешний адрес: " & strEmail)
          eCount = eCount + 1
        End If
      elseif (aclsContactList.Exists(strMBName) and aclsContactList.Exists(strDisName)) Then
        aclsContactList.Remove(strMBName)
        aclsContactList.Remove(strDisName)
      End If
    else
      log("Контакт в исключении, пропущен: " & strMBName & "(" & strDisName & ")")
      eCount = eCount + 1
    end if

'    wscript.echo strResult
    On Error Goto 0

  end if

  Set objField         = Nothing
  Set objContact       = Nothing
  Set objContainer     = Nothing
  strFirstName         = ""
  strLastName          = ""
  strAlias             = ""
  strEmail             = ""
  strlExDN             = ""
  strDisName           = ""
  strTDisName          = ""
  strMainDefault       = ""
  strProxy             = ""
  strMBName            = ""
  strMailbox1          = ""
  strDepartment        = ""
  strCompany           = ""
  strDescription       = ""
  strPhone             = ""
  strOffice            = ""
  strCustomAttribute1  = ""
  strCustomAttribute2  = ""
  strCustomAttribute3  = ""
  strCcn               = ""
  strCMail             = ""
  strClExDN            = ""
  strCgivenName        = ""
  strCsn               = ""
  strCDisName          = ""
  strCeAttr            = ""
  strCPhone            = ""
  strCOffice           = ""
  strCmAPIRecipient    = ""
  strCinternetEncoding = ""
Next

Log ("Создано   : " & iCount)
Log ("Пропущено : " & eCount)
Log ("Исправлено: " & xCount)
Log ("К удалению: " & aclsContactList.Count)
Log ("")

set objOU =GetObject("LDAP://ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt")
strTDisName = ""

On Error Resume Next : Err.Clear
For Each Item In aclsContactList
'  Set objContact = GetObject("LDAP://" & "cn=" & aclsContactList.Item(Item) & ",ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt")
'  Set objContact = GetObject("LDAP://" & "CN=" & UnEscape(Item) & ",OU=Contacts,OU=test,OU=P-house,DC=p-house,DC=pvt")
'  objContact.GetInfo
'  if Trim(LCase(CStr(objContact.Get("displayName")))) = Item then
'  If objContact.class = strObject Then            ' Каждый обьект класса "contact"
'    Log(aclsContactList.Item(Item) & " Удален!")
    Log(Item & " Удален!")
'    On Error Goto 0
'    objContact.DeleteObject(0)

    strTDisName = Item
    strTDisName = replace(strTDisName,"\","\\")
    strTDisName = replace(strTDisName,"/","\/")
    strTDisName = replace(strTDisName,"#","\#")
    strTDisName = replace(strTDisName,",","\,")

    objOU.Delete strObject, "CN=" & strTDisName
'    On Error Resume Next : Err.Clear
    aclsContactList.Remove(Item)
'  else
'    Log(Item & " - нужно удалить вручную.")
'  End If
'  Set objContact       = Nothing
   strTDisName=""
Next
On Error Goto 0

Set objOU = Nothing

Log ("Осталось  : " & aclsContactList.Count)
Log ("")
Log ("Всё!")

objSession.Logoff
Set objOu = Nothing
Set objFilter = Nothing
Set objAddrEntries= Nothing
Set objSession = Nothing
Set fso = Nothing
Set aclsContactList = Nothing
Set objField = Nothing

wscript.quit
