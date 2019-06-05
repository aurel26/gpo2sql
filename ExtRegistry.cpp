#include <Windows.h>
#include <msxml6.h>
#include <atlcomcli.h>
#include <stdio.h>
#include "gpo2sql.h"

extern HANDLE g_hHeap;
extern BOOL g_bSupportsAnsi;

BOOL
pExtPolicy (
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ LONG lPolicyId,
   _In_ IXMLDOMNode *pXmlNode
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXMLNodeList = NULL;

   long lLength;

   hr = pXmlNode->get_childNodes(&pXMLNodeList);
   hr = pXMLNodeList->get_length(&lLength);

   BSTR szName = NULL;
   BSTR szComment = NULL;
   BSTR szState = NULL;
   BSTR szCategory = NULL;

   for (long i = 0; i < lLength; i++)
   {
      IXMLDOMNode *pXmlSubNode;

      hr = pXMLNodeList->get_item(i, &pXmlSubNode);
      if (hr == S_OK)
      {
         BSTR strNodeName;

         hr = pXmlSubNode->get_baseName(&strNodeName);
         if (wcscmp(strNodeName, L"Name") == 0)
         {
            hr = pXmlSubNode->get_text(&szName);
         }
         else if (wcscmp(strNodeName, L"State") == 0)
         {
            hr = pXmlSubNode->get_text(&szState);
         }
         else if (wcscmp(strNodeName, L"Explain") == 0)
         {

         }
         else if (wcscmp(strNodeName, L"Supported") == 0)
         {

         }
         else if (wcscmp(strNodeName, L"Category") == 0)
         {
            hr = pXmlSubNode->get_text(&szCategory);
         }
         else if (wcscmp(strNodeName, L"Comment") == 0)
         {
            hr = pXmlSubNode->get_text(&szComment);
         }
         //
         // Value
         //
         else if (wcscmp(strNodeName, L"Text") == 0)
         {
            // Just test line
         }
         else if (wcscmp(strNodeName, L"EditText") == 0)
         {
            LPWSTR szName;
            LPWSTR szState;
            LPWSTR szValue;

            szName = XmlGetNodeValue(pXmlSubNode, L"./registry:Name");
            szState = XmlGetNodeValue(pXmlSubNode, L"./registry:State");
            szValue = XmlGetNodeValue(pXmlSubNode, L"./registry:Value");

            OutputWrite(
               lGpoId, szGpoScope, L"registry_value", 5,
               lPolicyId, L"EditText", szName, szState, szValue
            );

            _SafeStringFree(szName);
            _SafeStringFree(szState);
            _SafeStringFree(szValue);
         }
         else if (wcscmp(strNodeName, L"MultiText") == 0)
         {

         }
         else if (wcscmp(strNodeName, L"Numeric") == 0)
         {
            LPWSTR szName;
            LPWSTR szState;
            LPWSTR szValue;

            szName = XmlGetNodeValue(pXmlSubNode, L"./registry:Name");
            szState = XmlGetNodeValue(pXmlSubNode, L"./registry:State");
            szValue = XmlGetNodeValue(pXmlSubNode, L"./registry:Value");

            OutputWrite(
               lGpoId, szGpoScope, L"registry_value", 5,
               lPolicyId, L"Numeric", szName, szState, szValue
            );

            _SafeStringFree(szName);
            _SafeStringFree(szState);
            _SafeStringFree(szValue);
         }
         else if (wcscmp(strNodeName, L"DropDownList") == 0)
         {
            LPWSTR szName;
            LPWSTR szState;
            LPWSTR szValue;

            szName = XmlGetNodeValue(pXmlSubNode, L"./registry:Name");
            szState = XmlGetNodeValue(pXmlSubNode, L"./registry:State");
            szValue = XmlGetNodeValue(pXmlSubNode, L"./registry:Value/registry:Name");

            OutputWrite(
               lGpoId, szGpoScope, L"registry_value", 5,
               lPolicyId, L"DropDownList", szName, szState, szValue
            );

            _SafeStringFree(szName);
            _SafeStringFree(szState);
            _SafeStringFree(szValue);
         }
         else if (wcscmp(strNodeName, L"CheckBox") == 0)
         {
            LPWSTR szName;
            LPWSTR szState;

            szName = XmlGetNodeValue(pXmlSubNode, L"./registry:Name");
            szState = XmlGetNodeValue(pXmlSubNode, L"./registry:State");

            OutputWrite(
               lGpoId, szGpoScope, L"registry_value", 5,
               lPolicyId, L"CheckBox", szName, szState, NULL
            );

            _SafeStringFree(szName);
            _SafeStringFree(szState);
         }
         else if (wcscmp(strNodeName, L"ListBox") == 0)
         {

         }
         else
            Log(
               0, LOG_LEVEL_ERROR,
               L"ExtRegistry, unkwnon policy attribute ('%s').", strNodeName
            );

         _SafeStringFree(strNodeName);
         _SafeCOMRelease(pXmlSubNode);
      }
   }

   OutputWrite(
      lGpoId, szGpoScope, L"registry_policy", 5,
      lPolicyId, szName, szComment, szState, szCategory
   );

   _SafeStringFree(szName);
   _SafeStringFree(szComment);
   _SafeStringFree(szState);
   _SafeStringFree(szCategory);

   _SafeCOMRelease(pXMLNodeList);

   return TRUE;
}

BOOL
ExtProcessRegistry (
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXMLDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXmlNodeListPolicy = NULL;

   LONG lPolicyId = 1;
   long lLength;

   hr = pXMLDoc->setProperty((BSTR)L"SelectionNamespaces", CComVariant(L"xmlns:registry='http://www.microsoft.com/GroupPolicy/Settings/Registry'"));
   hr = pXmlNodeExtensionData->selectNodes((BSTR)L".//registry:Policy", &pXmlNodeListPolicy);
   hr = pXmlNodeListPolicy->get_length(&lLength);

   for (long i = 0; i < lLength; i++)
   {
      IXMLDOMNode *pXmlNodePolicy = NULL;

      hr = pXmlNodeListPolicy->get_item(i, &pXmlNodePolicy);

      if (hr == S_OK)
      {
         pExtPolicy(lGpoId, szGpoScope, lPolicyId, pXmlNodePolicy);
         lPolicyId++;
         _SafeCOMRelease(pXmlNodePolicy);
      }
   }

   _SafeCOMRelease(pXmlNodeListPolicy);

   return TRUE;
}