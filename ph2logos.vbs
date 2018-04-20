'Скрипт для импорта контактов из ПХ
'https://gist.github.com/rustavy/91027a71c363c5514dca
'
'v 1.2.3 07.06.17 13:42 by ZRO@mail.ru
'Неправильно изменялся аттрибут "title"
'v 1.2.2 15.05.17 14:57 by ZRO@mail.ru
'Добавил город
'v 1.2.1 01.04.15 14:55 by ZRO@mail.ru
'
'=========================================================================================
'Option Explicit

CONST strServer                             = "xhe.p-house.pvt"
CONST strMailbox                            = "gallg"
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
CONST CdoPR_HOME_ADDRESS_CITY               = &H3A59001E  'Home City
Const CdoPR_BUSINESS_ADDRESS_CITY           = &H3A27001E  'City
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
Const CdoPR_HIDE_FROM_ADDRESS_BOOK          = &H80B9000B  'True = Display
                                                          'False = Hide
Const PR_EMS_AB_PROXY_ADDRESSES             = &H800F101E
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
Dim sCurPath
Dim strFirstName, strLastName, strContactName, strEmail, strDisName, strMainDefault, strProxy, strMBName, strMailbox1
Dim strDepartment, strCompany, strDescription, strPhone, strOffice, strCity
Dim strCustomAttribute1, strCustomAttribute2, strCustomAttribute3
Dim strCcn, strCMail, strCgivenName, strCsn, strCDisName, strCeAttr, strCPhone, strCOffice, strCCity

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
'Set fso = CreateObject("Scripting.FileSystemObject")

if logg=1 or logg=3 then
  Set fso = CreateObject("Scripting.FileSystemObject")
  sCurPath = fso.GetParentFolderName(Wscript.ScriptFullName) & "\ph2logos.log"

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
    wscript.echo strTimeStamp & " " & sData
  end if
  strTimeStamp = ""
End Sub

'-------------------------------------------------------------'
'Important create OU=Contacts or change value for strContainer
'-------------------------------------------------------------'

set objOU =GetObject("LDAP://ou=press-house-Al,OU=logos,dc=logosgroup,dc=pvt")

iCount = 0

Set aclsContactList = CreateObject("Scripting.Dictionary")

'!!!!!ВНИМАНИЕ. Так делаем сейчас, т.е. добавляются только новые контакты.
' Собираем список имеющихся контактов
Log ("Собираем список имеющихся контактов...")
For Each objUser In objOU
  If objUser.class = strObject Then
    aclsContactList.Add objUser.cn, sAMAccountName ' Список выглядит как "cn" "sAMAccountName"
    iCount = iCount + 1
'    wscript.echo "x"
  End If
Next
Set objOU = Nothing
Log ("Найдено   : " & iCount)

'-------------------------------------------------------------'
'Задаем OU в которой будут создаваться контакты
'-------------------------------------------------------------'
Set objRootLDAP = GetObject("LDAP://rootDSE")
Set objContainer = GetObject("LDAP://ou=Press-House-Al,OU=logos,dc=logosgroup,dc=pvt")

'-------------------------------------------------------------'
'Сбрасывам переменные
'-------------------------------------------------------------'
iCount      = 0
eCount      = 0
xCount      = 0
strCcn      = ""
strCMail    = ""
strCsn      = ""
strCDisName = ""
strCeAttr   = ""
strmAPIRecipient = false
strinternetEncoding = 1310720
strCmAPIRecipient = ""
strCinternetEncoding = ""
strCCity    = ""

'-------------------------------------------------------------'
'Начинаем проверять
'-------------------------------------------------------------'
log("Начинаем проверять...")

'-------------------------------------------------------------'
'Для каждого контакта в адресной книге выполняем следующее
'-------------------------------------------------------------'
For Each objAddressEntry In objAddrEntries
'-------------------------------------------------------------'
'Если в поле legacyExchangeDN содержится /O=LOGOS-M/OU=P-HOUSE/cn=Recipients/cn="
'                                                      ^^^^^^^
'-------------------------------------------------------------'
  if mid(objAddressEntry.Address,15,7)="P-HOUSE" then 
    On Error Resume Next    

    '-------------------------------------------------------------'
    'Извлекаем данные контакта из адресной книги
    '-------------------------------------------------------------'
    strFirstName = Trim(objAddressEntry.Fields(CdoPR_GIVEN_NAME).Value)
    strLastName = Trim(objAddressEntry.Fields(CdoPR_SURNAME).Value)
'    strFirstName = Trim(CStr(objAddressEntry.Fields(CdoPR_GIVEN_NAME).Value))
'    strLastName = Trim(CStr(objAddressEntry.Fields(CdoPR_SURNAME).Value))
    strContactName = Trim(CStr(objAddressEntry.Name))

    Set objField = objAddressEntry.Fields(PR_EMS_AB_PROXY_ADDRESSES)
    
    For Each v In objField.Value   
      If Mid(v, 1, 4) = "SMTP" Then
        strEmail = Trim(Mid(v, 6))
        strMainDefault = Trim(v)
        strProxy = Trim(v)
      End If 
    Next
    
    strMBName           = Trim(mid(objAddressEntry.Address, instrrev(objAddressEntry.Address,"=")+1))
    strMailbox1         = Trim("/O=LOGOS-M/OU=P-HOUSE/cn=Recipients/cn=" & strMBName)
    strDisName          = Trim(objAddressEntry.Fields(CdoPR_DISPLAY_NAME).Value)
    strDepartment       = Trim(objAddressEntry.Fields(CdoPR_DEPARTMENT_NAME).Value)
    strCompany          = Trim(objAddressEntry.Fields(CdoPR_COMPANY_NAME).Value)
    strDescription      = Trim(objAddressEntry.Fields(CdoPR_TITLE).Value)
    strPhone            = Trim(CStr(objAddressEntry.Fields(CdoPR_OFFICE_TELEPHONE_NUMBER).Value))
    strOffice           = Trim(CStr(objAddressEntry.Fields(CdoPR_OFFICE_LOCATION).Value))
    strCity             = Trim(objAddressEntry.Fields(CdoPR_BUSINESS_ADDRESS_CITY).Value)
    strCustomAttribute1 = Trim(objAddressEntry.Fields(CdoPR_CUSTOM_ATTRIBUTE_1).Value)
    strCustomAttribute2 = Trim(objAddressEntry.Fields(CdoPR_CUSTOM_ATTRIBUTE_2).Value)
    strCustomAttribute3 = Trim(objAddressEntry.Fields(CdoPR_CUSTOM_ATTRIBUTE_3).Value)

'    On Error Goto 0

    If strEmail <> "stp@logosgroup.ru" and strEmail <> "stp@finans-media.ru" and strEmail <> "epak@mail.ru" Then ' Пропускаем исключения.
      If aclsContactList.Exists(strMBName) Then ' Если контакт есть в списке - пропускаем.
'        Set objContact = GetObject("LDAP://" & "cn=" & strMBName & ",ou=Contacts,ou=test,OU=p-house,dc=p-house,dc=pvt")
        Set objContact = GetObject("LDAP://" & "cn=" & strMBName & ",OU=Press-House-AL,OU=Logos,DC=logosgroup,DC=pvt")
'        On Error Resume Next : Err.Clear

        strCcn        = Trim(LCase(CStr(objContact.Get("cn"))))
        strCMail      = Trim(LCase(CStr(objContact.Get("Mail"))))
        strCgivenName = Trim(LCase(CStr(objContact.Get("givenName"))))
        strCsn        = Trim(LCase(CStr(objContact.Get("sn"))))
        strCDisName   = Trim(LCase(CStr(objContact.Get("displayName"))))
        strCeAttr     = Trim(LCase(CStr(objContact.Get("extensionAttribute1"))))
        strCPhone     = Trim(LCase(CStr(objContact.Get("telephoneNumber"))))
        strCOffice    = Trim(LCase(CStr(objContact.Get("physicalDeliveryOfficeName"))))
        strCCity      = Trim(LCase(CStr(objContact.Get("l"))))
        strCmAPIRecipient = Trim(LCase(CStr(objContact.Get("mAPIRecipient"))))
        strCinternetEncoding = Trim(LCase(CStr(objContact.Get("internetEncoding"))))

        if LCase(strMBName)<>strCcn Then
          log("Искали: `" & strMBName & "` Нашли: `" & strCcn & "`")
        end if

'           (strCeAttr<>LCase(strCustomAttribute1) and strCeAttr<>"прессхаус") or _
        If strCMail<>LCase(strEmail) or _
           (strCgivenName<>LCase(strFirstName) and strCgivenName<>"") or _
           (strCsn<>LCase(strLastName) and strCsn<>"") or _
           (strCDisName<>LCase(strDisName) and strCDisName<>"") or _
           strCeAttr<>LCase(strCustomAttribute1) or _
           strCPhone<>LCase(strPhone) or _
           strCmAPIRecipient<>LCase(strmAPIRecipient) or _
           strCinternetEncoding<>LCase(strinternetEncoding) or _
           strCCity<>LCase(strCity) or _
           strCOffice<>LCase(strOffice) Then

          log("Контакту нужны обновления: " & strEmail)
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
'          If strCeAttr<>LCase(strCustomAttribute1) and strCeAttr<>"'пх', ооо" then
          If strCeAttr<>LCase(strCustomAttribute1) then
            log("Контакту необходимо обновить атрибут: `" & strCeAttr & "`>`" & LCase(strCustomAttribute1) & "`")
          end if
          If strCPhone<>LCase(strPhone) then
            log("Контакту необходимо обновить тлф   : `" & strCPhone & "`>`" & LCase(strPhone) & "`")
          end if
          If strCOffice<>LCase(strOffice) then
            log("Контакту необходимо обновить офис  : `" & strCOffice & "`>`" & LCase(strOffice) & "`")
          end if
          If strCCity<>LCase(strCity) then
            log("Контакту необходимо обновить город  : `" & strCCity & "`>`" & LCase(strCity) & "`")
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
'-------------------
          if strCmAPIRecipient<>LCase(strmAPIRecipient) then
            log("Контакту необходимо обновить mAPIRecipient : `" & strCmAPIRecipient & "`>`" & LCase(strmAPIRecipient) & "`")
'            objContact.PutEx ADS_PROPERTY_UPDATE, "mAPIRecipient", strmAPIRecipient
            objContact.Put "mAPIRecipient", strmAPIRecipient
            log("mAPIRecipient : " & strmAPIRecipient)
          end if
          if strCinternetEncoding<>LCase(strinternetEncoding) then
            log("Контакту необходимо обновить internetEncoding : `" & strCinternetEncoding & "`>`" & LCase(strinternetEncoding) & "`")
'            objContact.PutEx ADS_PROPERTY_UPDATE, "internetEncoding", strinternetEncoding
            objContact.Put "internetEncoding", strinternetEncoding
            log("internetEncoding : " & strinternetEncoding)
          end if
'-------------------
          if strContactName <> "" then
            objContact.Put "mailNickname", UnEscape(strContactName)
            log("mailNickname     : " & strContactName)
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
            objContact.PutEx ADS_PROPERTY_UPDATE, "company", "'ПХ', ООО"
            log("company          : 'ПХ', ООО'")
          end if
          If strDescription <> "" Then
            objContact.Put "description", UnEscape(strDescription)
            objContact.Put "title", UnEscape(strDescription)
            log("description      : " & strDescription)
          else
            objContact.PutEx ADS_PROPERTY_UPDATE, "description", UnEscape("'ПХ', ООО")
            objContact.Put  "title", UnEscape("'ПХ', ООО")
            log("description      : 'ПХ', ООО")
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
          if strCity <> "" Then
            if strCCity <> "" Then
              objContact.Put "l", UnEscape(strCity)
            else
              objContact.Put "l", UnEscape(strCity)
            end if
            log("l                : " & strCity)
          else
            objContact.PutEx ADS_PROPERTY_CLEAR, "l", vbNullString
            log("l                : " & strCity)
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
            objContact.PutEx ADS_PROPERTY_CLEAR, "extensionAttribute1", vbNullString
'            objContact.PutEx ADS_PROPERTY_UPDATE, "extensionAttribute1", "ПрессХаус"
            log("Custom Attribute1: ''")
'            log("Custom Attribute1: 'ПрессХаус'")
          end if
          if strCustomAttribute2 <> "" then
            objContact.Put "extensionAttribute2", UnEscape(strCustomAttribute2)
            log("Custom Attribute2: " & strCustomAttribute2)
          end if
          if strCustomAttribute3 <> "" then
            objContact.Put "extensionAttribute3", UnEscape(strCustomAttribute3)
            log("Custom Attribute3: " & strCustomAttribute3)
          end if

          Err.Clear
'   On Error Goto 0
          objContact.SetInfo

          If (Err.number <> 0) Then
            log("Попытались записать данные - Что-то пошло не так...")
            log("Ошибка: " & Err.Number)
            Select Case Err.Number
            Case 424
              log("Не смогли изменить обьект.")
            Case &h80072032
              log("LDAP_INVALID_DN_SYNTAX This error occurs when a distinguished name used for the creation of objects contains invalid characters.")
            Case &h80071392
              log("Object with this name already exists.")
            Case &h80072030
              log("LDAP_NO_SUCH_OBJECT This error is similar to ADS_BAD_PATHNAME (0x80005008) - during the BIND process, an LDAP object path was passed from a non existing object.")
            Case &h8007200B
              log("The attribute syntax specified to the directory service is invalid.")
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
          log("Контакт существует, пропущен: " & strMBName)
          eCount = eCount + 1
        end if
        aclsContactList.Remove(strMBName)
        aclsContactList.Remove(strDisName)
'------------------------------------------------------------------------------
      Else                                      ' Если контакта нет в списке - создаем.
'        if instrrev(objAddressEntry.Address,"@presshouse.ru") <> 0 then ' Если у контакта адрес "@presshouse.ru"
        if instrrev(strEmail,"@presshouse.ru") <> 0 then ' Если у контакта адрес "@presshouse.ru"
          log("")
          log("Добавляем контакт: " & strMBName)
'          log("Добавляем контакт: " & strLastName & " " & strFirstName)
          Set objContact = objContainer.Create("Contact","cn=" & strMBName)
'          Set objContact = objContainer.Create("Contact","cn=" & strLastName & " " & strFirstName)
          if strFirstName <> "" then
            objContact.Put "givenName", strFirstName
            log("givenName        : " & strFirstName)
          end if
          if strLastName <> "" then
            objContact.Put "sn", strLastName
            log("sn               : " & strLastName)
          end if
          if strEmail <> "" then
            objContact.Put "Mail", strEmail
            log("Mail             : " & strEmail)
          end if
          if strProxy <> "" then
            objContact.Put "proxyAddresses", strProxy
            log("proxyAddresses   : " & strProxy)
          end if
          if strMainDefault <> "" then
            objContact.Put "targetAddress", strMainDefault
            log("targetAddress    : " & strMainDefault)
          end if
          if strMailbox1 <> "" then
            objContact.Put "legacyExchangeDN", strMailbox1
            log("legacyExchangeDN : " & strMailbox1)
          end if
'-------------------
'          if strCmAPIRecipient<>LCase(strmAPIRecipient) then
            objContact.Put "mAPIRecipient", strmAPIRecipient
            log("mAPIRecipient : " & strmAPIRecipient)
'          end if
'          if strCinternetEncoding<>LCase(strinternetEncoding) then
            objContact.Put "internetEncoding", strinternetEncoding
            log("internetEncoding : " & strinternetEncoding)
'          end if
'-------------------
          if strContactName <> "" then
            objContact.Put "mailNickname", strContactName
            log("mailNickname     : " & strContactName)
          end if
          if strContactName <> "" then
            objContact.Put "displayName", strContactName
            log("displayName      : " & strContactName)
          end if
          if strDepartment <> "" Then
            objContact.Put "department", strDepartment
            log("department       : " & strDepartment)
          end if
          If strCompany <> "" Then
            objContact.Put "company", strCompany
            log("company          : " & strCompany)
          else
            objContact.Put "company", "'ПХ', ООО"
            log("company          : 'ПХ', ООО")
          end if
          If strDescription <> "" Then
            objContact.Put "description", strDescription
            objContact.Put "title", strDescription
            log("description      : " & strDescription)
          else
            objContact.PutEx ADS_PROPERTY_UPDATE, "description", UnEscape("'ПХ', ООО")
            objContact.Put  "title", UnEscape("'ПХ', ООО")
            log("description      : 'ПХ', ООО")
          end if
          if strPhone <> "" Then
            objContact.Put "telephoneNumber", strPhone
            log("telephoneNumber  : " & strPhone)
          end if
          if strOffice <> "" Then
            objContact.Put "physicalDeliveryOfficeName", strOffice
            log("physicalDeliveryOfficeName: " & strOffice)
          end if
          if strCity <> "" Then
            objContact.Put "l", strCity
            log("l                : " & strCity)
          end if

'          objContact.SetInfo
          
          if strCustomAttribute1 <> "" then
            objContact.Put "extensionAttribute1", strCustomAttribute1
            log("Custom Attribute1: " & strCustomAttribute1)
          end if
          if strCustomAttribute2 <> "" then
            objContact.Put "extensionAttribute2", strCustomAttribute2
            log("Custom Attribute2: " & strCustomAttribute2)
          end if
          if strCustomAttribute3 <> "" then
            objContact.Put "extensionAttribute3", strCustomAttribute3
            log("Custom Attribute3: " & strCustomAttribute3)
          end if
          
'          On Error Resume Next : Err.Clear
'          If Err.Number = 0 Then
'            Err.Clear
'            objContact.MailEnable strProxy
'            If Err.Number <> 0 Then
'              log("Error Mail Enabling " & strMBName & ": " & Err.Description)
'            End If
'          Else
'            log("Error Configuring " & strMBName & ": " & Err.Description)
'          End If
'          On Error Goto 0
          
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
            Case &h8007200B
              log("The attribute syntax specified to the directory service is invalid.")
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
      End If
    else
      log("Контакт в исключении, пропущен: " & strMBName)
      eCount = eCount + 1
    end if

'    wscript.echo strResult
    On Error Goto 0

  end if

  Set objField         = Nothing
  Set objContact       = Nothing
  strFirstName         = ""
  strLastName          = ""
  strContactName       = ""
  strEmail             = ""
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
  strCity              = ""
  strCustomAttribute1  = ""
  strCustomAttribute2  = ""
  strCustomAttribute3  = ""
  strCcn               = ""
  strCMail             = ""
  strCgivenName        = ""
  strCsn               = ""
  strCDisName          = ""
  strCeAttr            = ""
  strCPhone            = ""
  strCOffice           = ""
  strCCity             = ""
  strCmAPIRecipient    = ""
  strCinternetEncoding = ""
Next

Log ("Создано   : " & iCount)
Log ("Пропущено : " & eCount)
Log ("Исправлено: " & xCount)
Log ("К удалению: " & aclsContactList.Count)
Log ("")

set objOU =GetObject("LDAP://ou=press-house-Al,OU=logos,dc=logosgroup,dc=pvt") 
strTDisName = ""

On Error Resume Next : Err.Clear
For Each Item In aclsContactList
    strTDisName = Item
    strTDisName = replace(strTDisName,"\","\\")
    strTDisName = replace(strTDisName,"/","\/")
    strTDisName = replace(strTDisName,"#","\#")
    strTDisName = replace(strTDisName,",","\,")
    Log(strTDisName & " Удален!")

    On Error Goto 0
    objOU.Delete strObject, "CN=" & strTDisName
    On Error Resume Next : Err.Clear
    aclsContactList.Remove(Item)

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
