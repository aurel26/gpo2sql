#include <Windows.h>
#include <msxml6.h>
#include <atlcomcli.h>
#include <stdio.h>
#include "gpo2sql.h"

extern HANDLE g_hHeap;
extern BOOL g_bSupportsAnsi;

BOOL
ExtProcessSecurity (
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXMLDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXmlNodeListExtension = NULL;

   long lLength;

   //XmlDebugShowNodes(pXmlNodeExtensionData, (LPWSTR)L"");

   hr = pXMLDoc->setProperty((BSTR)L"SelectionNamespaces", CComVariant((BSTR)L"xmlns:gen='http://www.microsoft.com/GroupPolicy/Settings' xmlns:security='http://www.microsoft.com/GroupPolicy/Settings/Security' xmlns:type='http://www.microsoft.com/GroupPolicy/Types'"));
   hr = pXmlNodeExtensionData->selectNodes((BSTR)L"./gen:Extension/*", &pXmlNodeListExtension);
   hr = pXmlNodeListExtension->get_length(&lLength);

   for (long i = 0; i < lLength; i++)
   {
      IXMLDOMNode *pXmlNodeExtension = NULL;
      BSTR strNodeName;

      hr = pXmlNodeListExtension->get_item(i, &pXmlNodeExtension);
      hr = pXmlNodeExtension->get_baseName(&strNodeName);

      if (wcscmp(strNodeName, L"Blocked") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"SecurityOptions") == 0)
      {
         LPWSTR  szName, szValue, szDisplayName, szDisplayString;

         szName = XmlGetNodeValue(pXmlNodeExtension, L"./security:KeyName");
         if (szName == NULL)
         {
            szName = XmlGetNodeValue(pXmlNodeExtension, L"./security:SystemAccessPolicyName");
            if (szName == NULL)
            {
               Log(
                  0, LOG_LEVEL_WARNING,
                  L"ExtRegistry, no name."
               );
            }
         }

         szValue= XmlGetNodeValue(pXmlNodeExtension, L"./security:SettingNumber");
         if (szName == NULL)
         {
            szValue = XmlGetNodeValue(pXmlNodeExtension, L"./security:string");
            if (szName == NULL)
            {
               szValue = XmlGetNodeValue(pXmlNodeExtension, L"./security:boolean");
               if (szName == NULL)
               {
                  Log(
                     0, LOG_LEVEL_WARNING,
                     L"ExtRegistry, no value."
                  );
               }
            }
         }

         szDisplayName = XmlGetNodeValue(pXmlNodeExtension, L"./security:Display/security:Name");
         szDisplayString = XmlGetNodeValue(pXmlNodeExtension, L"./security:Display/security:DisplayString");

         OutputWrite(
            lGpoId, szGpoScope, L"security_option", 4,
            szName, szValue, szDisplayName, szDisplayString
         );

         _SafeStringFree(szName);
         _SafeStringFree(szValue);
         _SafeStringFree(szDisplayName);
         _SafeStringFree(szDisplayString);
      }
      else if (wcscmp(strNodeName, L"RestrictedGroups") == 0)
      {
         LPWSTR szGroupNameSID, szGroupNameName, szMemberOfSID, szMemberOfName;

         szGroupNameSID = XmlGetNodeValue(pXmlNodeExtension, L"./security:GroupName/type:SID");
         szGroupNameName = XmlGetNodeValue(pXmlNodeExtension, L"./security:GroupName/type:Name");
         szMemberOfSID = XmlGetNodeValue(pXmlNodeExtension, L"./security:Memberof/type:SID");
         szMemberOfName = XmlGetNodeValue(pXmlNodeExtension, L"./security:Memberof/type:Name");

         OutputWrite(
            lGpoId, szGpoScope, L"security_restricted_group", 4,
            szGroupNameSID, szGroupNameName, szMemberOfSID, szMemberOfName
         );

         _SafeStringFree(szGroupNameSID);
         _SafeStringFree(szGroupNameName);
         _SafeStringFree(szMemberOfSID);
         _SafeStringFree(szMemberOfName);
      }
      else if (wcscmp(strNodeName, L"UserRightsAssignment") == 0)
      {
         IXMLDOMNodeList *pXmlNodeListMember= NULL;
         long lLength;
         LPWSTR szName;

         szName = XmlGetNodeValue(pXmlNodeExtension, L"./security:Name");
         hr = pXmlNodeExtension->selectNodes((BSTR)L"./security:Member", &pXmlNodeListMember);
         hr = pXmlNodeListMember->get_length(&lLength);

         if (lLength == 0)
         {
            OutputWrite(
               lGpoId, szGpoScope, L"security_userright_assignment", 3,
               szName, NULL, NULL
            );
         }
         else
         {
            for (long i = 0; i < lLength; i++)
            {
               IXMLDOMNode *pXmlNodeMember = NULL;
               LPWSTR szMemberSID, szMemberName;

               hr = pXmlNodeListMember->get_item(i, &pXmlNodeMember);
               szMemberSID = XmlGetNodeValue(pXmlNodeMember, L"./type:SID");
               szMemberName = XmlGetNodeValue(pXmlNodeMember, L"./type:Name");

               OutputWrite(
                  lGpoId, szGpoScope, L"security_userright_assignment", 3,
                  szName, szMemberSID, szMemberName
               );

               _SafeStringFree(szMemberSID);
               _SafeStringFree(szMemberName);
               _SafeCOMRelease(pXmlNodeMember);
            }
         }

         _SafeStringFree(szName);
         _SafeCOMRelease(pXmlNodeListMember);
      }
      else if (wcscmp(strNodeName, L"EventLog") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Account") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"SystemServices") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Audit") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"File") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Registry") == 0)
      {

      }
      else
      {
         wprintf(L"%s\n", strNodeName);
      }

      _SafeStringFree(strNodeName);
      _SafeCOMRelease(pXmlNodeExtension);
   }

   _SafeCOMRelease(pXmlNodeListExtension);

   return TRUE;
}
