#include "Output.h"

//
// Log levels
//
#define LOG_LEVEL_NONE        0   // Tracing is not on
#define LOG_LEVEL_CRITICAL    1   // Abnormal exit or termination
#define LOG_LEVEL_ERROR       2   // Severe errors that need logging
#define LOG_LEVEL_WARNING     3   // Warnings such as allocation failure
#define LOG_LEVEL_INFORMATION 4   // Includes non-error cases(e.g.,Entry-Exit)
#define LOG_LEVEL_VERBOSE     5   // Detailed traces from intermediate steps
#define LOG_LEVEL_VERYVERBOSE 6

//
// Colors VT100
//
#define COLOR_RED          (g_bSupportsAnsi) ? L"\x1b[38;5;1m" : L""
#define COLOR_GREEN        (g_bSupportsAnsi) ? L"\x1b[1;32m" : L""
#define COLOR_YELLOW       (g_bSupportsAnsi) ? L"\x1b[1;33m" : L""
#define COLOR_BRIGHT_RED   (g_bSupportsAnsi) ? L"\x1b[38;5;9m" : L""
#define COLOR_MAGENTA      (g_bSupportsAnsi) ? L"\x1b[1;35m" : L""
#define COLOR_CYAN         (g_bSupportsAnsi) ? L"\x1b[1;36m" : L""
#define COLOR_WHITE        (g_bSupportsAnsi) ? L"\x1b[1;37m" : L""
#define COLOR_RESET        (g_bSupportsAnsi) ? L"\x1b[0m" : L""

//
// Macros
//
#define _HeapAlloc(x) HeapAlloc(g_hHeap, HEAP_ZERO_MEMORY, (x))
#define _SafeHeapRelease(x) { if (NULL != x) { HeapFree(g_hHeap, 0, x); x = NULL; } }
#define _SafeCOMRelease(x) { if (NULL != x) { x->Release(); x = NULL; } }
#define _SafeStringFree(x) { if (NULL != x) { SysFreeString(x); x = NULL; } }

//
// Structures
//
typedef struct _GLOBAL_CONFIG
{
   LPCWSTR szDatabaseName;
   LPCWSTR szInputDirectory;
   LPCWSTR szOutputDirectory;
} GLOBAL_CONFIG, *PGLOBAL_CONFIG;

typedef struct _GPO
{
   GUID Identifier;
   LPWSTR Domain;
   LPWSTR Name;
   LPWSTR CreatedTime;
   LPWSTR ModifiedTime;
   BOOL FilterDataAvailable;
   LPWSTR FilterName;
   LPWSTR FilterDescription;
   int Computer_VersionDirectory;
   int Computer_VersionSysvol;
   BOOL Computer_Enabled;
   int User_VersionDirectory;
   int User_VersionSysvol;
   BOOL User_Enabled;
} GPO, *PGPO;

typedef struct _LINKS_TO
{
   LPWSTR SOMName;
   LPWSTR SOMPath;
   BOOL Enabled;
   BOOL NoOverride;
} LINKS_TO, *PLINKS_TO;

//
// XmlParser.cpp
//
BOOL
XmlParseFile(
   _In_ PLONG plGpoId,
   _In_z_ LPCWSTR szFileName
);

LPWSTR
XmlGetNodeValue(
   _In_ IXMLDOMNode *pXmlNode,
   _In_z_ LPCWSTR szPath
);

LPWSTR
XmlGetNodeAttribute(
   _In_ IXMLDOMNode *pXmlNode,
   _In_z_ LPCWSTR szAttributeName
);

BOOL
XmlFillFromChildNodes(
   _In_ IXMLDOMNode *pXmlNode,
   _In_ DWORD dwChildNodesCount,
   _In_ LPCWSTR *szChildNodesNames,
   _Out_ LPWSTR **szOutput
);

BOOL
XmlListFromChildNodes(
   _In_ IXMLDOMNode *pXmlNode,
   _In_ LPCWSTR szChildNodePath,
   _Out_writes_(dwOutputSize) LPWSTR szOutput,
   _In_ DWORD dwOutputSize
);

BOOL
XmlFillFromAttributes(
   _In_ IXMLDOMNode *pXmlNode,
   _In_ DWORD dwAttributesCount,
   _In_ LPCWSTR *szAttributesNames,
   _Out_ LPWSTR **szOutput
);

BOOL
XmlDebugShowNodes(
   _In_ IXMLDOMNode *pXmlNode,
   _Inout_updates_z_(dwPathSize) LPWSTR szPath
);

//
// Util.cpp
//
VOID
Log(
   _In_ DWORD dwIndent,
   _In_ DWORD dwLevel,
   _In_z_ LPCWSTR szFormat,
   ...
);

VOID
RemoveSpecialChars(
   _In_z_ LPWSTR szString
);

LPSTR
LPWSTRtoLPSTR(
   _In_opt_z_ LPWSTR szToConvert
);

VOID
StringDuplicate(
   _In_z_ LPWSTR szInput,
   _Out_ LPWSTR *szOutput
);

BOOL
WriteTextFile(
   _In_ HANDLE hFile,
   _In_z_ LPCSTR szFormat,
   ...
);

//
// Output
//
BOOL
OutputInitialize(
);

BOOL
OutputWrite(
   _In_ LONG lGpoId,
   _In_opt_z_ LPCWSTR szType,
   _In_z_ LPCWSTR szBase,
   _In_ DWORD dwParamCount,
   ...
);

BOOL
OutputClose(
);