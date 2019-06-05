#include <Windows.h>
#include <stdio.h>
#include <stdarg.h>
#include "gpo2sql.h"

#define MSG_MAX_SIZE       8192

extern HANDLE g_hLogFile;
extern HANDLE g_hHeap;
extern BOOL g_bSupportsAnsi;

VOID
Log (
   _In_ DWORD dwIndent,
   _In_ DWORD dwLevel,
   _In_z_ LPCWSTR szFormat,
   ...
)
{
   int r;

   WCHAR szMessage[MSG_MAX_SIZE];
   SYSTEMTIME st;

   va_list argptr;
   va_start(argptr, szFormat);

   GetLocalTime(&st);

   r = vswprintf_s(szMessage, MSG_MAX_SIZE, szFormat, argptr);
   if (r == -1)
   {
      return;
   }

   for (DWORD i = 0; i < dwIndent; i++)
   {
      wprintf(L"  ");
   }

   if (dwLevel <= LOG_LEVEL_INFORMATION)
      wprintf(L"%s\n", szMessage);

   /*
   if (dwLevel <= LOG_LEVEL_VERBOSE)
   {
      DWORD dwDataSize, dwDataWritten;
      WCHAR szLine[INFO_MAX_SIZE];

      swprintf_s(
         szLine, INFO_MAX_SIZE,
         L"%04u/%02u/%02u - %02u:%02u:%02u.%03u\t%s\r\n",
         st.wYear, st.wMonth, st.wDay,
         st.wHour, st.wMinute, st.wSecond, st.wMilliseconds,
         szMessage
      );

      dwDataSize = (DWORD)wcsnlen_s(szLine, INFO_MAX_SIZE);
      WriteFile(g_hLogFile, szLine, dwDataSize * sizeof(WCHAR), &dwDataWritten, NULL);
   }
   */
}

VOID
RemoveSpecialChars(
   _In_z_ LPWSTR szString
)
{
   // Remove \r \n \t
   if (szString)
   {
      while (*szString)
      {
         if (*szString == 0x0a)
            *szString = 0x20;
         else if (*szString == 0x0d)
            *szString = 0x20;
         else if (*szString == 0x09)
            *szString = 0x20;
         szString++;
      }
   }
}

LPSTR
LPWSTRtoLPSTR(
   _In_opt_z_ LPWSTR szToConvert
)
{
   LPSTR szResult;
   int iSize;

   if (szToConvert == NULL)
      return NULL;

   iSize = WideCharToMultiByte(
      CP_ACP,
      0,
      szToConvert,
      -1,
      NULL, 0,
      NULL, NULL
   );

   if (iSize == 0)
      goto Fail;

   szResult = (LPSTR)HeapAlloc(g_hHeap, HEAP_ZERO_MEMORY, iSize + 1);

   if (szResult == NULL)
      goto Fail;

   iSize = WideCharToMultiByte(
      CP_ACP,
      0,
      szToConvert,
      -1,
      szResult, iSize,
      NULL, NULL
   );

   if (iSize == 0)
   {
      _SafeHeapRelease(szResult);
      goto Fail;
   }

   return szResult;

Fail:
   Log(
      0, LOG_LEVEL_ERROR,
      L"LPWSTRtoLPSTR('%s') failed.", szToConvert
   );

   return NULL;
}

VOID
StringDuplicate (
   _In_z_ LPWSTR szInput,
   _Out_ LPWSTR *szOutput
)
{
   size_t InputSize;

   if ((szInput == NULL) || (szOutput == NULL))
      return;

   InputSize = wcslen(szInput);
   *szOutput = (LPWSTR)_HeapAlloc((InputSize + 1) * sizeof(WCHAR));
   if (*szOutput == NULL)
      return;

   memcpy(*szOutput, szInput, InputSize * sizeof(WCHAR));
}

BOOL
WriteTextFile (
   _In_ HANDLE hFile,
   _In_z_ LPCSTR szFormat,
   ...
)
{
   BOOL bResult;
   DWORD dwDataSize, dwDataWritten;
   CHAR szMessage[MSG_MAX_SIZE];

   va_list argptr;
   va_start(argptr, szFormat);

   vsprintf_s(szMessage, MSG_MAX_SIZE, szFormat, argptr);

   dwDataSize = (DWORD)strnlen_s(szMessage, MSG_MAX_SIZE);
   bResult = WriteFile(hFile, szMessage, dwDataSize, &dwDataWritten, NULL);
   if (bResult == FALSE)
   {
      Log(
         0, LOG_LEVEL_ERROR,
         L"%sWriteFile() failed%s (error %u).",
         COLOR_RED, COLOR_RESET, GetLastError()
      );
      return FALSE;
   }

   return TRUE;
}
