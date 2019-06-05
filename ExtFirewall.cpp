#include <Windows.h>
#include <msxml6.h>
#include <atlcomcli.h>
#include <stdio.h>
#include "gpo2sql.h"

#define MAX_PROFILE_SIZE 22

extern HANDLE g_hHeap;
extern BOOL g_bSupportsAnsi;

BOOL
pParseProfile (
   _In_ LONG lGpoId,
   _In_ IXMLDOMNode *pXmlNodeExtensionData,
   _In_z_ LPCWSTR szXmlNodePath,
   _In_z_ LPCWSTR szProfileName
)
{
   HRESULT hr;

   IXMLDOMNode *pXmlNodeListProfile = NULL;

   hr = pXmlNodeExtensionData->selectSingleNode((BSTR)szXmlNodePath, &pXmlNodeListProfile);

   if (hr == S_OK)
   {
      LPWSTR szEnableFirewall;
      LPWSTR szDoNotAllowExceptions;
      LPWSTR szDefaultInboundAction;
      LPWSTR szDefaultOutboundAction;
      LPWSTR szDisableNotifications;
      LPWSTR szAllowLocalPolicyMerge;
      LPWSTR szAllowLocalIPsecPolicyMerge;
      LPWSTR szLogFilePath;
      LPWSTR szLogFileSize;
      LPWSTR szLogDroppedPackets;
      LPWSTR szLogSuccessfulConnections;

      szEnableFirewall = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:EnableFirewall/fw:Value");
      szDefaultInboundAction = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:DefaultInboundAction/fw:Value");
      szDefaultOutboundAction = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:DefaultOutboundAction/fw:Value");
      szDoNotAllowExceptions = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:DoNotAllowExceptions/fw:Value");
      szDisableNotifications = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:DisableNotifications/fw:Value");
      szAllowLocalPolicyMerge = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:AllowLocalPolicyMerge/fw:Value");
      szAllowLocalIPsecPolicyMerge = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:AllowLocalIPsecPolicyMerge/fw:Value");
      szLogFilePath = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:LogFilePath/fw:Value");
      szLogFileSize = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:LogFileSize/fw:Value");
      szLogDroppedPackets = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:LogDroppedPackets/fw:Value");
      szLogSuccessfulConnections = XmlGetNodeValue(pXmlNodeListProfile, L"./fw:LogSuccessfulConnections/fw:Value");

      OutputWrite(
         lGpoId, NULL,
         L"firewall_profile", 12,
         szProfileName,
         szEnableFirewall,
         szDefaultInboundAction,
         szDefaultOutboundAction,
         szDoNotAllowExceptions,
         szDisableNotifications,
         szAllowLocalPolicyMerge,
         szAllowLocalIPsecPolicyMerge,
         szLogFilePath,
         szLogFileSize,
         szLogDroppedPackets,
         szLogSuccessfulConnections
      );

       _SafeStringFree(szEnableFirewall);
       _SafeStringFree(szDoNotAllowExceptions);
       _SafeStringFree(szDefaultInboundAction);
       _SafeStringFree(szDefaultOutboundAction);
       _SafeStringFree(szDisableNotifications);
       _SafeStringFree(szAllowLocalPolicyMerge);
       _SafeStringFree(szAllowLocalIPsecPolicyMerge);
       _SafeStringFree(szLogFilePath);
       _SafeStringFree(szLogFileSize);
       _SafeStringFree(szLogDroppedPackets);
       _SafeStringFree(szLogSuccessfulConnections);

      _SafeCOMRelease(pXmlNodeListProfile);
   }

   return TRUE;
}

BOOL
pParseRules (
   _In_ LONG lGpoId,
   _In_ IXMLDOMNode *pXmlNodeExtensionData,
   _In_z_ LPCWSTR szXmlNodePath,
   _In_z_ LPCWSTR szType
)
{
   HRESULT hr;

   IXMLDOMNodeList *pXmlNodeRuleCollection = NULL;

   long lLength;

   hr = pXmlNodeExtensionData->selectNodes((BSTR)szXmlNodePath, &pXmlNodeRuleCollection);
   if (hr == S_OK)
   {
      hr = pXmlNodeRuleCollection->get_length(&lLength);

      for (long i = 0; i < lLength; i++)
      {
         IXMLDOMNode *pXmlNodeRule = NULL;

         hr = pXmlNodeRuleCollection->get_item(i, &pXmlNodeRule);

         if (hr == S_OK)
         {
            WCHAR szProfile[MAX_PROFILE_SIZE];

            LPWSTR szVersion, szAction, szName, szDesc;
            LPWSTR szActive, szEmbedCtxt, szApp, szSvc, szDir;
            LPWSTR szProtocol, szICMP4, szICMP6, szLPort, szRPort;
            LPWSTR szLA4, szLA6, szRA4, szRA6, szSecurity;

            LPCWSTR szInput[] = {
               L"Version", L"Action", L"Name",L"Desc",
               L"Active", L"EmbedCtxt", L"App", L"Svc", L"Dir",
               L"Protocol", L"ICMP4", L"ICMP6", L"LPort", L"RPort",
               L"LA4", L"LA6", L"RA4", L"RA6", L"Security"
            };

            LPWSTR* szOutput[] = {
               &szVersion, &szAction, &szName, &szDesc,
               &szActive, &szEmbedCtxt, &szApp, &szSvc, &szDir,
               &szProtocol, &szICMP4, &szICMP6, &szLPort, &szRPort,
               &szLA4, &szLA6, &szRA4, &szRA6, &szSecurity
            };

            XmlListFromChildNodes(pXmlNodeRule, L"./fw:Profile", szProfile, MAX_PROFILE_SIZE);
            XmlFillFromChildNodes(pXmlNodeRule, 19, szInput, szOutput);

            OutputWrite(
               lGpoId, NULL,
               L"firewall_rule", 21,
               szType,
               szProfile,
               szVersion,
               szAction,
               szName,
               szDesc,
               szActive,
               szEmbedCtxt,
               szApp,
               szSvc,
               szDir,
               szProtocol,
               szICMP4,
               szICMP6,
               szLPort,
               szRPort,
               szLA4,
               szLA6,
               szRA4,
               szRA6,
               szSecurity
            );

            _SafeStringFree(szVersion);
            _SafeStringFree(szAction);
            _SafeStringFree(szName);
            _SafeStringFree(szDesc);
            _SafeStringFree(szActive);
            _SafeStringFree(szEmbedCtxt);
            _SafeStringFree(szApp);
            _SafeStringFree(szSvc);
            _SafeStringFree(szDir);
            _SafeStringFree(szProtocol);
            _SafeStringFree(szICMP4);
            _SafeStringFree(szICMP6);
            _SafeStringFree(szLPort);
            _SafeStringFree(szRPort);
            _SafeStringFree(szLA4);
            _SafeStringFree(szLA6);
            _SafeStringFree(szRA4);
            _SafeStringFree(szRA6);
            _SafeStringFree(szSecurity);

            _SafeCOMRelease(pXmlNodeRule);
         }
      }
      _SafeCOMRelease(pXmlNodeRuleCollection);
   }

   return TRUE;
}

BOOL
ExtFirewall (
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXMLDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
)
{
   HRESULT hr;

   hr = pXMLDoc->setProperty((BSTR)L"SelectionNamespaces", CComVariant(L"xmlns:fw='http://www.microsoft.com/GroupPolicy/Settings/WindowsFirewall'"));

   //
   // Profiles
   //
   pParseProfile(lGpoId, pXmlNodeExtensionData, L".//fw:PublicProfile", L"Public");
   pParseProfile(lGpoId, pXmlNodeExtensionData, L".//fw:PrivateProfile", L"Private");
   pParseProfile(lGpoId, pXmlNodeExtensionData, L".//fw:DomainProfile", L"Domain");

   pParseRules(lGpoId, pXmlNodeExtensionData, L".//fw:InboundFirewallRules", L"Inbound");
   pParseRules(lGpoId, pXmlNodeExtensionData, L".//fw:OutboundFirewallRules", L"Outbound");

   return TRUE;
}