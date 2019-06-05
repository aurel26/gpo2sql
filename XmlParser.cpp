#include <Windows.h>
#include <msxml6.h>
#include <atlcomcli.h>
#include "gpo2sql.h"
#include "XmlExt.h"

extern HANDLE g_hHeap;
extern BOOL g_bSupportsAnsi;

BOOL
pXmlProcessExtensionData (
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXmlDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
)
{
   HRESULT hr;

   IXMLDOMNode *pXmlNodeName = NULL;
   IXMLDOMNode *pXmlNodeError = NULL;
   BSTR strNodeName;

   hr = pXmlDoc->setProperty((BSTR)L"SelectionNamespaces", CComVariant(L"xmlns:gen='http://www.microsoft.com/GroupPolicy/Settings'"));

   hr = pXmlNodeExtensionData->selectSingleNode((BSTR)L"./gen:Name", &pXmlNodeName);
   hr = pXmlNodeName->get_text(&strNodeName);

   hr = pXmlNodeExtensionData->selectSingleNode((BSTR)L"./gen:Error", &pXmlNodeError);
   if (hr == S_OK)
   {
      //hr = pXmlNodeName->get_text(&strNodeText);
      //_SafeStringFree(strNodeText);
      _SafeCOMRelease(pXmlNodeError);
   }
   else
   {
      OutputWrite(
         lGpoId, szGpoScope, L"extension", 1,
         strNodeName
      );

      if (wcscmp(strNodeName, L"Registry") == 0)
      {
         ExtProcessRegistry(lGpoId, szGpoScope, pXmlDoc, pXmlNodeExtensionData);
      }
      else if (wcscmp(strNodeName, L"Security") == 0)
      {
         ExtProcessSecurity(lGpoId, szGpoScope, pXmlDoc, pXmlNodeExtensionData);
      }
      else if (wcscmp(strNodeName, L"Scripts") == 0)
      {
         ExtProcessScript(lGpoId, szGpoScope, pXmlDoc, pXmlNodeExtensionData);
      }
      else if (wcscmp(strNodeName, L"Public Key") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Advanced Audit Configuration") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Local Users and Groups") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Software Installation") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Remote Installation") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Scheduled Tasks") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Name Resolution Policy") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Power Options") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Regional Options") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Internet Options") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Internet Explorer Maintenance") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Drive Maps") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Folders") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Services") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Shortcuts") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Files") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Application Control Policies") == 0)
      {
         ExtAppControl(lGpoId, szGpoScope, pXmlDoc, pXmlNodeExtensionData);
      }
      else if (wcscmp(strNodeName, L"Software Restriction") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Environment Variables") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Folder Redirection") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Start Menu") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"WLanSvc Networks") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"LanSvc Networks") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Policy-based QoS") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"IP security") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Printers") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Deployed Printer Connections Policy") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Windows Registry") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Windows Firewall") == 0)
      {
      ExtFirewall(lGpoId, szGpoScope, pXmlDoc, pXmlNodeExtensionData);
      }
      else
      {
         Log(
            0, LOG_LEVEL_WARNING,
            L"Unknown XML tag (%s). Please report.", strNodeName
         );
      }
   }

   _SafeStringFree(strNodeName);
   _SafeCOMRelease(pXmlNodeName);

   return TRUE;
}

BOOL
pXmlParseLinksTo (
   _In_ LONG lGpoId,
   _In_ IXMLDOMNode *pXmlNode
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXMLNodeList = NULL;

   LINKS_TO LinksTo = { 0 };

   long lLength;

   hr = pXmlNode->get_childNodes(&pXMLNodeList);
   hr = pXMLNodeList->get_length(&lLength);

   for (long i = 0; i < lLength; i++)
   {
      IXMLDOMNode *pXmlNodeValue = NULL;
      BSTR strNodeName;
      BSTR strNodeValue;

      hr = pXMLNodeList->get_item(i, &pXmlNodeValue);
      hr = pXmlNodeValue->get_nodeName(&strNodeName);
      hr = pXmlNodeValue->get_text(&strNodeValue);

      if (wcscmp(strNodeName, L"SOMName") == 0)
      {
         StringDuplicate(strNodeValue, &LinksTo.SOMName);
      }
      else if (wcscmp(strNodeName, L"SOMPath") == 0)
      {
         StringDuplicate(strNodeValue, &LinksTo.SOMPath);
      }
      else if (wcscmp(strNodeName, L"Enabled") == 0)
      {
         if (_wcsicmp(strNodeValue, L"true") == 0)
            LinksTo.Enabled = TRUE;
      }
      else if (wcscmp(strNodeName, L"NoOverride") == 0)
      {
         if (_wcsicmp(strNodeValue, L"true") == 0)
            LinksTo.NoOverride = TRUE;
      }
      else
      {
         fwprintf_s(stderr, L"Unknwon property: %s\n", strNodeName);
      }

      _SafeStringFree(strNodeName);
      _SafeStringFree(strNodeValue);
      _SafeCOMRelease(pXmlNodeValue);
   }

   OutputWrite(
      lGpoId, NULL, L"linksto", 4,
      LinksTo.SOMName, LinksTo.SOMPath, LinksTo.Enabled, LinksTo.NoOverride
   );

   _SafeHeapRelease(LinksTo.SOMName);
   _SafeHeapRelease(LinksTo.SOMPath);

   _SafeCOMRelease(pXMLNodeList);

   return TRUE;
}

BOOL
pXmlParseNodes (
   _In_ LONG lGpoId,
   _In_ IXMLDOMDocument2 *pXmlDoc,
   _In_ IXMLDOMNode *pXmlNode,
   _In_z_ LPCWSTR szPath,
   _Out_ PGPO pGpo
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXMLNodeList = NULL;

   size_t SizePath;
   long lLength;

   BSTR strNodeName;
   BSTR strNodeValue;
   LPWSTR szNewPath;

   hr = pXmlNode->get_nodeName(&strNodeName);
   hr = pXmlNode->get_text(&strNodeValue);

   SizePath = wcslen(szPath) + wcslen(strNodeName) + 2;
   szNewPath = (LPWSTR)_HeapAlloc(SizePath * sizeof(WCHAR));
   swprintf_s(szNewPath, SizePath, L"%s/%s", szPath, strNodeName);

   hr = pXmlNode->get_childNodes(&pXMLNodeList);
   hr = pXMLNodeList->get_length(&lLength);

   if (_wcsicmp(szNewPath, L"/GPO/Identifier/Identifier") == 0)
   {
      if (wcslen(strNodeValue) == 38)
      {
         hr = CLSIDFromString(strNodeValue, &pGpo->Identifier);
      }
   }
   else if (_wcsicmp(szNewPath, L"/GPO/Identifier/Domain") == 0)
   {
      StringDuplicate(strNodeValue, &pGpo->Domain);
   }
   else if (_wcsicmp(szNewPath, L"/GPO/Name") == 0)
   {
      StringDuplicate(strNodeValue, &pGpo->Name);
   }
   else if (_wcsicmp(szNewPath, L"/GPO/CreatedTime") == 0)
   {
      StringDuplicate(strNodeValue, &pGpo->CreatedTime);
   }
   else if (_wcsicmp(szNewPath, L"/GPO/ModifiedTime") == 0)
   {
      StringDuplicate(strNodeValue, &pGpo->ModifiedTime);
   }
   else if (_wcsicmp(szNewPath, L"/GPO/FilterDataAvailable") == 0)
   {
      if (_wcsicmp(strNodeValue, L"true") == 0)
         pGpo->FilterDataAvailable = TRUE;
   }
   else if (_wcsicmp(szNewPath, L"/GPO/FilterName") == 0)
   {
      StringDuplicate(strNodeValue, &pGpo->FilterName);
   }
   else if (_wcsicmp(szNewPath, L"/GPO/FilterDescription") == 0)
   {
      StringDuplicate(strNodeValue, &pGpo->FilterDescription);
   }
   //
   // Computer
   //
   else if (_wcsicmp(szNewPath, L"/GPO/Computer/VersionDirectory") == 0)
   {
      int i, v;

      i = swscanf_s(strNodeValue, L"%d", &v);
      if (i == 1)
         pGpo->Computer_VersionDirectory = v;
   }
   else if (_wcsicmp(szNewPath, L"/GPO/Computer/VersionSysvol") == 0)
   {
      int i, v;

      i = swscanf_s(strNodeValue, L"%d", &v);
      if (i == 1)
         pGpo->Computer_VersionSysvol = v;
   }
   else if (_wcsicmp(szNewPath, L"/GPO/Computer/Enabled") == 0)
   {
      if (_wcsicmp(strNodeValue, L"true") == 0)
         pGpo->Computer_Enabled = TRUE;
   }
   //
   // User
   //
   else if (_wcsicmp(szNewPath, L"/GPO/User/VersionDirectory") == 0)
   {
      int i, v;

      i = swscanf_s(strNodeValue, L"%d", &v);
      if (i == 1)
         pGpo->User_VersionDirectory = v;
   }
   else if (_wcsicmp(szNewPath, L"/GPO/User/VersionSysvol") == 0)
   {
      int i, v;

      i = swscanf_s(strNodeValue, L"%d", &v);
      if (i == 1)
         pGpo->User_VersionSysvol = v;
   }
   else if (_wcsicmp(szNewPath, L"/GPO/User/Enabled") == 0)
   {
      if (_wcsicmp(strNodeValue, L"true") == 0)
         pGpo->User_Enabled = TRUE;
   }
   //
   // Security
   //
   else if (_wcsicmp(szNewPath, L"/GPO/SecurityDescriptor") == 0)
   {
      // Do nothing
   }
   else if (_wcsicmp(szNewPath, L"/GPO/LinksTo") == 0)
   {
      pXmlParseLinksTo(lGpoId, pXmlNode);
   }
   //
   // Extensions
   //
   else if (_wcsicmp(szNewPath, L"/GPO/Computer/ExtensionData") == 0)
   {
      pXmlProcessExtensionData(lGpoId, L"Computer", pXmlDoc, pXmlNode);
   }
   else if (_wcsicmp(szNewPath, L"/GPO/User/ExtensionData") == 0)
   {
      pXmlProcessExtensionData(lGpoId, L"User", pXmlDoc, pXmlNode);
   }
   else
   {
      //wprintf_s(L"%s\n", szNewPath);

      for (long i = 0; i < lLength; i++)
      {
         IXMLDOMNode *pXmlSubNode;

         hr = pXMLNodeList->get_item(i, &pXmlSubNode);
         pXmlParseNodes(lGpoId, pXmlDoc, pXmlSubNode, szNewPath, pGpo);
         _SafeCOMRelease(pXmlSubNode);
      }
   }

   _SafeStringFree(strNodeName);
   _SafeStringFree(strNodeValue);
   _SafeCOMRelease(pXMLNodeList);

   _SafeHeapRelease(szNewPath);

   return TRUE;
}

BOOL
pXmlParseXmlNode (
   _In_ LONG lGpoId,
   _In_ IXMLDOMDocument2 *pXmlDoc,
   _In_ IXMLDOMNode *pXmlNodeGpo
)
{
   GPO Gpo = { 0 };
   pXmlParseNodes(lGpoId, pXmlDoc, pXmlNodeGpo, (LPWSTR)L"", &Gpo);

   OutputWrite(
      lGpoId, NULL, L"gpo", 14,
      &Gpo.Identifier, Gpo.Domain, Gpo.Name,
      Gpo.CreatedTime, Gpo.ModifiedTime,
      Gpo.FilterDataAvailable, Gpo.FilterName, Gpo.FilterDescription,
      Gpo.Computer_VersionDirectory, Gpo.Computer_VersionSysvol, Gpo.Computer_Enabled,
      Gpo.User_VersionDirectory, Gpo.User_VersionSysvol, Gpo.User_Enabled
   );

   _SafeHeapRelease(Gpo.Domain);
   _SafeHeapRelease(Gpo.Name);
   _SafeHeapRelease(Gpo.CreatedTime);
   _SafeHeapRelease(Gpo.ModifiedTime);
   _SafeHeapRelease(Gpo.FilterName);
   _SafeHeapRelease(Gpo.FilterDescription);

   return TRUE;
}

BOOL
XmlParseFile (
   _In_ PLONG plGpoId,
   _In_z_ LPCWSTR szFileName
)
{
   VARIANT_BOOL bSuccess = false;

   HRESULT hr;

   IXMLDOMDocument2 *pXMLDoc = NULL;

   //
   // GPO data
   //
   Log(
      0, LOG_LEVEL_INFORMATION,
      L"%sProcess%s %s'%s'%s.",
      COLOR_CYAN, COLOR_RESET,
      COLOR_WHITE, szFileName, COLOR_RESET
   );

   hr = CoCreateInstance(CLSID_DOMDocument60, NULL, CLSCTX_INPROC_SERVER, IID_IXMLDOMDocument2, (void**)&pXMLDoc);
   if ((hr != S_OK) || (pXMLDoc == NULL))
   {
      Log(
         0, LOG_LEVEL_CRITICAL,
         L"Unable to create XML object (error 0x%08x).", hr
      );
      return NULL;
   }

   hr = pXMLDoc->put_async(VARIANT_FALSE);

   //
   // Load XML file
   //
   hr = pXMLDoc->load(CComVariant(szFileName), &bSuccess);

   if ((hr != S_OK) || (bSuccess == FALSE))
   {
      IXMLDOMParseError *pXmlParseError = NULL;
      BSTR strError;

      hr = pXMLDoc->get_parseError(&pXmlParseError);
      hr = pXmlParseError->get_reason(&strError);

      RemoveSpecialChars(strError);
      Log(
         0, LOG_LEVEL_CRITICAL,
         L"Unable to parse XML (%s).", strError
      );

      _SafeCOMRelease(pXmlParseError);
      _SafeCOMRelease(pXMLDoc);

      return FALSE;
   }

   hr = pXMLDoc->setProperty((BSTR)L"SelectionNamespaces", CComVariant(L"xmlns:gen='http://www.microsoft.com/GroupPolicy/Settings'"));

   //
   // Test 1: one GPO in file
   //
   {
      IXMLDOMNode *pXMLNodeGpo = NULL;

      hr = pXMLDoc->selectSingleNode((BSTR)L"/gen:GPO", &pXMLNodeGpo);

      if (hr == S_OK)
      {
         pXmlParseXmlNode(*plGpoId, pXMLDoc, pXMLNodeGpo);
      }
   }

   //
   // Test 2: multiple GPO in file (report)
   //
   {
      IXMLDOMNodeList *pXmlGpoNodeList = NULL;

      hr = pXMLDoc->selectNodes((BSTR)L"/report/gen:GPO", &pXmlGpoNodeList);

      if (hr == S_OK)
      {
         long lLength;

         hr = pXmlGpoNodeList->get_length(&lLength);

         for (long i = 0; i < lLength; i++)
         {
            IXMLDOMNode *pXMLNodeGpo = NULL;

            hr = pXmlGpoNodeList->get_item(i, &pXMLNodeGpo);
            if (hr == S_OK)
            {
               pXmlParseXmlNode(*plGpoId, pXMLDoc, pXMLNodeGpo);
               (*plGpoId)++;
               _SafeCOMRelease(pXMLNodeGpo);
            }
         }
         _SafeCOMRelease(pXmlGpoNodeList);
      }
   }

   _SafeCOMRelease(pXMLDoc);

   return TRUE;
   /*
   for (long i = 0; i < lLength; i++)
   {
      IXMLDOMNode *pXmlNodeValue = NULL;
      BSTR strNodeName = NULL;

      hr = pXMLNodeListGpo->get_item(i, &pXmlNodeValue);
      hr = pXmlNodeValue->get_nodeName(&strNodeName);

      if (wcscmp(strNodeName, L"Identifier") == 0)
      {
         IXMLDOMNodeList *pXMLNodeListIdentifier = NULL;
         long lLengthIdentifier;

         hr = pXmlNodeValue->get_childNodes(&pXMLNodeListIdentifier);
         hr = pXMLNodeListIdentifier->get_length(&lLengthIdentifier);
      }
      else if (wcscmp(strNodeName, L"Name") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"IncludeComments") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"CreatedTime") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"ModifiedTime") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"ReadTime") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"SecurityDescriptor") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"FilterDataAvailable") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"FilterName") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"FilterDescription") == 0)
      {

      }
      else if (wcscmp(strNodeName, L"Computer") == 0)
      {
         xmlProcessGpoSettings(pXmlNodeValue);
      }
      else if (wcscmp(strNodeName, L"User") == 0)
      {
         xmlProcessGpoSettings(pXmlNodeValue);
      }
      else if (wcscmp(strNodeName, L"LinksTo") == 0)
      {

      }
      else
      {
         Sleep(1);
      }

      _SafeStringFree(strNodeName);
      _SafeCOMRelease(pXmlNodeValue);
   }
   */
}

LPWSTR
XmlGetNodeValue (
   _In_ IXMLDOMNode *pXmlNode,
   _In_z_ LPCWSTR szPath
)
{
   HRESULT hr;
   IXMLDOMNode *pXmlNodeValue = NULL;

   hr = pXmlNode->selectSingleNode((BSTR)szPath, &pXmlNodeValue);
   if (hr == S_OK)
   {
      BSTR strNodeText;

      pXmlNodeValue->get_text(&strNodeText);
      _SafeCOMRelease(pXmlNodeValue);
      return strNodeText;
   }
   else
      return NULL;
}

LPWSTR
XmlGetNodeAttribute (
   _In_ IXMLDOMNode *pXmlNode,
   _In_z_ LPCWSTR szAttributeName
)
{
   HRESULT hr;
   IXMLDOMNamedNodeMap *pXmlAttributeMap;

   hr = pXmlNode->get_attributes(&pXmlAttributeMap);
   if (hr == S_OK)
   {
      IXMLDOMNode *pXmlNodeAttribute = NULL;

      hr = pXmlAttributeMap->getNamedItem((BSTR)szAttributeName, &pXmlNodeAttribute);
      if (hr == S_OK)
      {
         BSTR strName;

         hr = pXmlNodeAttribute->get_text(&strName);

         _SafeCOMRelease(pXmlNodeAttribute);
         _SafeCOMRelease(pXmlAttributeMap);
         return strName;
      }
      _SafeCOMRelease(pXmlAttributeMap);
      return NULL;
   }
   else
      return NULL;
}

BOOL
XmlFillFromChildNodes (
   _In_ IXMLDOMNode *pXmlNode,
   _In_ DWORD dwChildNodesCount,
   _In_ LPCWSTR *szChildNodesNames,
   _Out_ LPWSTR **szOutput
)
{
   HRESULT hr;
   IXMLDOMNodeList  *pXmlChildNodeList;

   for (DWORD j = 0; j < dwChildNodesCount; j++)
   {
      *szOutput[j] = NULL;
   }

   hr = pXmlNode->get_childNodes(&pXmlChildNodeList);
   if (hr == S_OK)
   {
      long lLength;

      hr = pXmlChildNodeList->get_length(&lLength);

      for (long i = 0; i < lLength; i++)
      {
         IXMLDOMNode *pXmlNode;

         hr = pXmlChildNodeList->get_item(i, &pXmlNode);
         if (hr == S_OK)
         {
            BSTR szNodeName;

            hr = pXmlNode->get_baseName(&szNodeName);
            if (hr == S_OK)
            {
               for (DWORD j = 0; j < dwChildNodesCount; j++)
               {
                  if (wcscmp(szNodeName, szChildNodesNames[j]) == 0)
                  {
                     BSTR szValue;

                     hr = pXmlNode->get_text(&szValue);
                     if (hr == S_OK)
                     {
                        *szOutput[j] = szValue;
                     }
                     break;
                  }
               }
               _SafeStringFree(szNodeName);
            }
            _SafeCOMRelease(pXmlNode);
         }
      }
      _SafeCOMRelease(pXmlChildNodeList);
   }

   return TRUE;
}

BOOL
XmlListFromChildNodes (
   _In_ IXMLDOMNode *pXmlNode,
   _In_ LPCWSTR szChildNodePath,
   _Out_writes_(dwOutputSize) LPWSTR szOutput,
   _In_ DWORD dwOutputSize
)
{
   BOOL bFirst = TRUE;
   HRESULT hr;
   IXMLDOMNodeList  *pXmlChildNodeList;

   if (dwOutputSize == 0)
      return FALSE;

   szOutput[0] = 0;

   hr = pXmlNode->selectNodes((BSTR)szChildNodePath, &pXmlChildNodeList);
   if (hr == S_OK)
   {
      long lLength;

      hr = pXmlChildNodeList->get_length(&lLength);

      for (long i = 0; i < lLength; i++)
      {
         IXMLDOMNode *pXmlNode;

         hr = pXmlChildNodeList->get_item(i, &pXmlNode);
         if (hr == S_OK)
         {
            BSTR szValue;

            hr = pXmlNode->get_text(&szValue);
            if (hr == S_OK)
            {
               if (bFirst == FALSE)
               {
                  wcscat_s(szOutput, dwOutputSize, L",");
               }
               else
               {
                  bFirst = FALSE;
               }
               wcscat_s(szOutput, dwOutputSize, szValue);
               _SafeStringFree(szValue);
            }
            _SafeCOMRelease(pXmlNode);
         }
      }
      _SafeCOMRelease(pXmlChildNodeList);
   }

   return TRUE;
}

BOOL
XmlFillFromAttributes (
   _In_ IXMLDOMNode *pXmlNode,
   _In_ DWORD dwAttributesCount,
   _In_ LPCWSTR *szAttributesNames,
   _Out_ LPWSTR **szOutput
)
{
   HRESULT hr;
   IXMLDOMNamedNodeMap *pXmlAttributeMap;

   for (DWORD j = 0; j < dwAttributesCount; j++)
   {
      *szOutput[j] = NULL;
   }

   hr = pXmlNode->get_attributes(&pXmlAttributeMap);
   if (hr == S_OK)
   {
      long lLength;

      hr = pXmlAttributeMap->get_length(&lLength);

      for (long i = 0; i < lLength; i++)
      {
         IXMLDOMNode *pXmlNode;

         hr = pXmlAttributeMap->get_item(i, &pXmlNode);
         if (hr == S_OK)
         {
            BSTR szAttributeName;

            hr = pXmlNode->get_baseName(&szAttributeName);
            if (hr == S_OK)
            {
               for (DWORD j = 0; j < dwAttributesCount; j++)
               {
                  if (wcscmp(szAttributeName, szAttributesNames[j]) == 0)
                  {
                     BSTR szValue;

                     hr = pXmlNode->get_text(&szValue);
                     if (hr == S_OK)
                     {
                        *szOutput[j] = szValue;
                     }
                     break;
                  }
               }
               _SafeStringFree(szAttributeName);
            }
            _SafeCOMRelease(pXmlNode);
         }
      }
      _SafeCOMRelease(pXmlAttributeMap);
   }

   return TRUE;
}

BOOL
XmlDebugShowNodes (
   _In_ IXMLDOMNode *pXmlNode,
   _Inout_updates_z_(dwPathSize) LPWSTR szPath
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXMLNodeList = NULL;

   size_t SizePath;
   long lLength;

   BSTR strNodeName;
   BSTR strNodeValue;
   LPWSTR szNewPath;

   hr = pXmlNode->get_nodeName(&strNodeName);
   hr = pXmlNode->get_text(&strNodeValue);

   SizePath = wcslen(szPath) + wcslen(strNodeName) + 2;
   szNewPath = (LPWSTR)_HeapAlloc(SizePath * sizeof(WCHAR));
   swprintf_s(szNewPath, SizePath, L"%s/%s", szPath, strNodeName);

   hr = pXmlNode->get_childNodes(&pXMLNodeList);
   hr = pXMLNodeList->get_length(&lLength);

   wprintf_s(L"%s\n", szNewPath);

   for (long i = 0; i < lLength; i++)
   {
      IXMLDOMNode *pXmlSubNode;

      hr = pXMLNodeList->get_item(i, &pXmlSubNode);
      XmlDebugShowNodes(pXmlSubNode, szNewPath);
      _SafeCOMRelease(pXmlSubNode);
   }

   _SafeStringFree(strNodeName);
   _SafeStringFree(strNodeValue);
   _SafeCOMRelease(pXMLNodeList);

   _SafeHeapRelease(szNewPath);

   return TRUE;
}