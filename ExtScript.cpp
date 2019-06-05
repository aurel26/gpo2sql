#include <Windows.h>
#include <msxml6.h>
#include <atlcomcli.h>
#include <stdio.h>
#include "gpo2sql.h"

extern HANDLE g_hHeap;
extern BOOL g_bSupportsAnsi;

BOOL
ExtProcessScript (
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXMLDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXmlNodeListScript = NULL;

   long lLength;

   hr = pXMLDoc->setProperty((BSTR)L"SelectionNamespaces", CComVariant(L"xmlns:scripts='http://www.microsoft.com/GroupPolicy/Settings/Scripts'"));
   hr = pXmlNodeExtensionData->selectNodes((BSTR)L".//scripts:Script", &pXmlNodeListScript);
   hr = pXmlNodeListScript->get_length(&lLength);

   for (long i = 0; i < lLength; i++)
   {
      IXMLDOMNode *pXmlNodeScript = NULL;

      hr = pXmlNodeListScript->get_item(i, &pXmlNodeScript);

      if (hr == S_OK)
      {
         LPWSTR szCommand, szParameters, szType, szOrder;

         szCommand = XmlGetNodeValue(pXmlNodeScript, L"./scripts:Command");
         szParameters = XmlGetNodeValue(pXmlNodeScript, L"./scripts:Parameters");
         szType = XmlGetNodeValue(pXmlNodeScript, L"./scripts:Type");
         szOrder = XmlGetNodeValue(pXmlNodeScript, L"./scripts:Order");

         OutputWrite(
            lGpoId, szGpoScope, L"script", 4,
            szCommand, szParameters, szType, szOrder
         );

         _SafeStringFree(szCommand);
         _SafeStringFree(szParameters);
         _SafeStringFree(szType);
         _SafeStringFree(szOrder);

         _SafeCOMRelease(pXmlNodeScript);
      }
   }

   _SafeCOMRelease(pXmlNodeListScript);

   return TRUE;
}