//
// ExtRegistry.cpp
//
BOOL
ExtProcessRegistry(
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXMLDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
);

//
// ExtSecurity.cpp
//
BOOL
ExtProcessSecurity(
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXMLDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
);

//
// ExtAppControl.cpp
//
BOOL
ExtAppControl(
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXMLDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
);

//
// ExtScript.cpp
//
BOOL
ExtProcessScript(
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXMLDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
);

//
// ExtFirewall.cpp
//
BOOL
ExtFirewall(
   _In_ LONG lGpoId,
   _In_z_ LPCWSTR szGpoScope,
   _In_ IXMLDOMDocument2 *pXMLDoc,
   _In_ IXMLDOMNode *pXmlNodeExtensionData
);