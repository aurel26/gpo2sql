typedef enum _OUTPUT_TYPE
{
   TYPE_INTEGER,
   TYPE_STRING,
   TYPE_STRING_TO_BOOL,
   TYPE_STRING_TO_INTEGER,
   TYPE_GUID,
   TYPE_DATE,
   TYPE_BOOL
} OUTPUT_TYPE;

#define OUTPUT_OPTION_ADD_ID_GPO          1
#define OUTPUT_OPTION_ADD_SCOPE           2

typedef struct _OUTPUT_ENTRY
{
   OUTPUT_TYPE Type;
   LPCWSTR szName;
   DWORD dwSize;
} OUTPUT_ENTRY, *POUTPUT_ENTRY;

typedef struct _OUTPUT_BASE
{
   LPCWSTR szName;
   DWORD dwOptions;
   DWORD dwEntriesCount;
   OUTPUT_ENTRY *pEntries;
   HANDLE hFile;
} OUTPUT_BASE, *POUTPUT_BASE;