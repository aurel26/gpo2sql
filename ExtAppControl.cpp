#include <Windows.h>
#include <msxml6.h>
#include <atlcomcli.h>
#include <stdio.h>
#include "gpo2sql.h"

#define MAX_VALUE_SIZE        2048

extern HANDLE g_hHeap;
extern BOOL g_bSupportsAnsi;

BOOL
pParseFilePathRule (
   _In_ LONG lGpoId,
   _In_ DWORD dwIdRules,
   _In_ DWORD dwIdRule,
   _In_ IXMLDOMNode *pXmlRuleNode
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXMLFilePathList = NULL;

   hr = pXmlRuleNode->selectNodes((BSTR)L"./app:Conditions/app:FilePath", &pXMLFilePathList);

   if (hr == S_OK)
   {
      long lLength;
      DWORD dwIdCondition = 1;

      hr = pXMLFilePathList->get_length(&lLength);

      for (long i = 0; i < lLength; i++)
      {
         IXMLDOMNode *pXmlFilePathNode;

         hr = pXMLFilePathList->get_item(i, &pXmlFilePathNode);
         if (hr == S_OK)
         {
            LPWSTR szPath;

            szPath = XmlGetNodeAttribute(pXmlFilePathNode, L"Path");
            if (szPath != NULL)
            {
               OutputWrite(
                  lGpoId, NULL, L"appctl_condition", 4,
                  dwIdRules, dwIdRule, dwIdCondition, szPath
               );

               _SafeStringFree(szPath);
            }
            _SafeCOMRelease(pXmlFilePathNode);
         }
         dwIdCondition++;
      }
      _SafeCOMRelease(pXMLFilePathList);

   }

   return TRUE;
}

BOOL
pParseFilePublisherRule (
   _In_ LONG lGpoId,
   _In_ DWORD dwIdRules,
   _In_ DWORD dwIdRule,
   _In_ IXMLDOMNode *pXmlRuleNode
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXMLFilePublisherList = NULL;

   hr = pXmlRuleNode->selectNodes((BSTR)L"./app:Conditions/app:FilePublisher", &pXMLFilePublisherList);

   if (hr == S_OK)
   {
      long lLength;
      DWORD dwIdCondition = 1;

      hr = pXMLFilePublisherList->get_length(&lLength);

      for (long i = 0; i < lLength; i++)
      {
         IXMLDOMNode *pXmlPublisherNode;

         hr = pXMLFilePublisherList->get_item(i, &pXmlPublisherNode);
         if (hr == S_OK)
         {
            WCHAR szValue[MAX_VALUE_SIZE];

            LPWSTR szPublisherName;
            LPWSTR szProductName;
            LPWSTR szBinaryName;

            szPublisherName = XmlGetNodeAttribute(pXmlPublisherNode, L"PublisherName");
            szProductName = XmlGetNodeAttribute(pXmlPublisherNode, L"ProductName");
            szBinaryName = XmlGetNodeAttribute(pXmlPublisherNode, L"BinaryName");

            swprintf_s(
               szValue, MAX_VALUE_SIZE,
               L"PublisherName=\"%s\" ProductName=\"%s\" BinaryName=\"%s\"",
               szPublisherName, szProductName, szBinaryName
            );

            OutputWrite(
               lGpoId, NULL, L"appctl_condition", 4,
               dwIdRules, dwIdRule, dwIdCondition, szValue
            );

            _SafeStringFree(szPublisherName);
            _SafeStringFree(szProductName);
            _SafeStringFree(szBinaryName);
            _SafeCOMRelease(pXmlPublisherNode);
         }
         dwIdCondition++;
      }
      _SafeCOMRelease(pXMLFilePublisherList);

   }

   return TRUE;
}

BOOL
pParseFileHashRule (
   _In_ LONG lGpoId,
   _In_ DWORD dwIdRules,
   _In_ DWORD dwIdRule,
   _In_ IXMLDOMNode *pXmlRuleNode
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXMLFileHashList = NULL;

   hr = pXmlRuleNode->selectNodes((BSTR)L"./app:Conditions/app:FileHash", &pXMLFileHashList);

   if (hr == S_OK)
   {
      long lLength;
      DWORD dwIdCondition = 1;

      hr = pXMLFileHashList->get_length(&lLength);

      for (long i = 0; i < lLength; i++)
      {
         IXMLDOMNode *pXmlFileHashNode;

         hr = pXMLFileHashList->get_item(i, &pXmlFileHashNode);
         if (hr == S_OK)
         {
            IXMLDOMNode *pXmlFileInfoNode;

            WCHAR szValue[MAX_VALUE_SIZE];
            LPWSTR szType;

            szType = XmlGetNodeAttribute(pXmlFileHashNode, L"Type");

            hr = pXmlFileHashNode->selectSingleNode((BSTR)L"./app:FileInformation", &pXmlFileInfoNode);
            if (hr == S_OK)
            {
               LPWSTR szName;
               LPWSTR szLength;

               szName = XmlGetNodeAttribute(pXmlFileInfoNode, L"Name");
               szLength = XmlGetNodeAttribute(pXmlFileInfoNode, L"Length");

               swprintf_s(
                  szValue, MAX_VALUE_SIZE,
                  L"Type=\"%s\" Name=\"%s\" Length=\"%s\"",
                  szType, szName, szLength
               );

               _SafeStringFree(szName);
               _SafeStringFree(szLength);
               _SafeCOMRelease(pXmlFileInfoNode);
            }
            else
            {
               swprintf_s(
                  szValue, MAX_VALUE_SIZE,
                  L"Type=\"%s\"",
                  szType
               );
            }

            OutputWrite(
               lGpoId, NULL, L"appctl_condition", 4,
               dwIdRules, dwIdRule, dwIdCondition, szValue
            );

            _SafeStringFree(szType);
            _SafeCOMRelease(pXmlFileHashNode);
         }
         dwIdCondition++;
      }
      _SafeCOMRelease(pXMLFileHashList);

   }

   return TRUE;
}

BOOL
ExtAppControl (
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXMLDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXmlNodeListRuleCollection = NULL;

   long lLength;

   DWORD dwIdRules = 1;

   hr = pXMLDoc->setProperty((BSTR)L"SelectionNamespaces", CComVariant(L"xmlns:app='http://www.microsoft.com/GroupPolicy/Settings/SRPV2'"));
   hr = pXmlNodeExtensionData->selectNodes((BSTR)L".//app:RuleCollection", &pXmlNodeListRuleCollection);
   hr = pXmlNodeListRuleCollection->get_length(&lLength);

   for (long i = 0; i < lLength; i++)
   {
      IXMLDOMNode *pXmlNodeRuleCollection = NULL;

      hr = pXmlNodeListRuleCollection->get_item(i, &pXmlNodeRuleCollection);

      if (hr == S_OK)
      {
         IXMLDOMNodeList *pXMLRuleList = NULL;

         LPWSTR szType;
         LPWSTR szEnforcementMode;

         szType = XmlGetNodeAttribute(pXmlNodeRuleCollection, L"Type");
         szEnforcementMode = XmlGetNodeValue(pXmlNodeRuleCollection, L"./app:EnforcementMode/app:Mode");

         OutputWrite(
            lGpoId, szGpoScope, L"appctl_rule_collection", 3,
            dwIdRules, szType, szEnforcementMode
         );

         _SafeStringFree(szType);
         _SafeStringFree(szEnforcementMode);

         hr = pXmlNodeRuleCollection->get_childNodes(&pXMLRuleList);
         if (hr == S_OK)
         {
            long lLengthRule;
            DWORD dwIdRule = 1;

            hr = pXMLRuleList->get_length(&lLengthRule);

            for (long j = 0; j < lLengthRule; j++)
            {
               IXMLDOMNode *pXmlRuleNode;

               hr = pXMLRuleList->get_item(j, &pXmlRuleNode);
               if (hr == S_OK)
               {
                  BSTR szRuleName;

                  hr = pXmlRuleNode->get_baseName(&szRuleName);
                  if (hr == S_OK)
                  {
                     LPCWSTR szType = NULL;

                     if (wcscmp(szRuleName, L"EnforcementMode") == 0)
                     {
                        // Do nothing
                     }
                     else if (wcscmp(szRuleName, L"FilePathRule") == 0)
                     {
                        szType = L"FilePath";
                        pParseFilePathRule(lGpoId, dwIdRules, dwIdRule, pXmlRuleNode);
                     }
                     else if (wcscmp(szRuleName, L"FilePublisherRule") == 0)
                     {
                        szType = L"FilePublisher";
                        pParseFilePublisherRule(lGpoId, dwIdRules, dwIdRule, pXmlRuleNode);
                     }
                     else if (wcscmp(szRuleName, L"FileHashRule") == 0)
                     {
                        szType = L"FileHashRule";
                        pParseFileHashRule(lGpoId, dwIdRules, dwIdRule, pXmlRuleNode);
                     }
                     else
                     {
                        Log(
                           0, LOG_LEVEL_WARNING,
                           L"Unknown XML tag (%s). Please report.", szRuleName
                        );
                     }

                     if (szType != NULL)
                     {
                        LPWSTR szName, szDescription, szUserOrGroupSid, szAction;
                        LPCWSTR szInput[] = { L"Name", L"Description", L"UserOrGroupSid", L"Action" };
                        LPWSTR* szOutput[] = { &szName, &szDescription, &szUserOrGroupSid, &szAction };

                        XmlFillFromAttributes(pXmlRuleNode, 4, szInput, szOutput);

                        OutputWrite(
                           lGpoId, NULL, L"appctl_rule", 7,
                           dwIdRules, dwIdRule, szType, szName, szDescription, szUserOrGroupSid, szAction
                        );
                     }

                     _SafeStringFree(szRuleName);
                  }
                  _SafeCOMRelease(pXmlRuleNode);
               }
               dwIdRule++;
            }
            _SafeCOMRelease(pXMLRuleList);
         }
         _SafeCOMRelease(pXmlNodeRuleCollection);
      }
      dwIdRules++;
   }

   _SafeCOMRelease(pXmlNodeListRuleCollection);

   return TRUE;
}