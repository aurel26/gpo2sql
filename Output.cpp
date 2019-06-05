#include <Windows.h>
#include <stdio.h>
#include <stdarg.h>
#include "gpo2sql.h"
#include "Output_Bases.h"

extern GLOBAL_CONFIG g_GlobalConfig;

#define STR_FMT_VERSION    "14.0"
#define STR_FMT_SQLSORT    "SQL_Latin1_General_CP1_CI_AS"
#define STR_FMT_EMPTY      "\"\""

BOOL
pOutputWriteFile (
   _In_ POUTPUT_BASE pOutputBase,
   _In_z_ LPWSTR szText,
   _Out_opt_ PDWORD pdwSize
)
{
   size_t sizeText;
   DWORD dwBytesWritten;

   sizeText = wcslen(szText);
   WriteFile(pOutputBase->hFile, szText, sizeText * sizeof(WCHAR), &dwBytesWritten, NULL);

   if (pdwSize != NULL)
   {
      if (*pdwSize < sizeText)
         *pdwSize = sizeText;
   }

   return TRUE;
}

BOOL
pCreateSqlFile (
   _In_ HANDLE hSqlFile,
   _In_ HANDLE hFmtFile,
   _In_ POUTPUT_BASE pOutputBase
)
{
   DWORD dwColumnCount, dwColumnIndex = 1;
   LPCSTR szDelimiter;

   szDelimiter = R"("\t\0")";

   dwColumnCount = pOutputBase->dwEntriesCount;
   if (pOutputBase->dwOptions & OUTPUT_OPTION_ADD_ID_GPO)
      dwColumnCount++;
   if (pOutputBase->dwOptions & OUTPUT_OPTION_ADD_SCOPE)
      dwColumnCount++;

   WriteTextFile(hSqlFile, "CREATE TABLE[dbo].[%S]\n(\n", pOutputBase->szName);

   WriteTextFile(hFmtFile, "%s\r\n", STR_FMT_VERSION);
   WriteTextFile(hFmtFile, "%u\r\n", dwColumnCount);

   if (pOutputBase->dwOptions & OUTPUT_OPTION_ADD_ID_GPO)
   {
      WriteTextFile(hFmtFile, "%u SQLNCHAR 0 0 %s %u %s %s\r\n", dwColumnIndex, szDelimiter, dwColumnIndex, "id_gpo", STR_FMT_EMPTY);
      WriteTextFile(hSqlFile, "   [id_gpo] int,\n");
      dwColumnIndex++;
   }

   if (pOutputBase->dwOptions & OUTPUT_OPTION_ADD_SCOPE)
   {
      WriteTextFile(hFmtFile, "%u SQLNCHAR 0 0 %s %u %s %s\r\n", dwColumnIndex, szDelimiter, dwColumnIndex, "scope", STR_FMT_SQLSORT);
      WriteTextFile(hSqlFile, "   [scope] nvarchar(8),\n");
      dwColumnIndex++;
   }

   for (DWORD i = 0; i < pOutputBase->dwEntriesCount; i++)
   {
      LPCSTR szComma;

      if (i == (pOutputBase->dwEntriesCount - 1))
      {
         szComma = "";
         szDelimiter = R"("\r\0\n\0")";               // Last entry
      }
      else
      {
         szComma = ",";
         szDelimiter = R"("\t\0")";
      }

      switch (pOutputBase->pEntries[i].Type)
      {
         case TYPE_STRING_TO_INTEGER:
         case TYPE_INTEGER:
         {
            WriteTextFile(hFmtFile, "%u SQLNCHAR 0 0 %s %u %S %s\r\n", dwColumnIndex, szDelimiter, dwColumnIndex, pOutputBase->pEntries[i].szName, STR_FMT_EMPTY);
            WriteTextFile(hSqlFile, "   [%S] int DEFAULT(NULL)%s\n", pOutputBase->pEntries[i].szName, szComma);
            break;
         }

         case TYPE_STRING:
         case TYPE_DATE:
         {
            WriteTextFile(hFmtFile, "%u SQLNCHAR 0 0 %s %u %S %s\r\n", dwColumnIndex, szDelimiter, dwColumnIndex, pOutputBase->pEntries[i].szName, STR_FMT_SQLSORT);
            WriteTextFile(hSqlFile, "   [%S] nvarchar(%u) DEFAULT(NULL)%s\n", pOutputBase->pEntries[i].szName, pOutputBase->pEntries[i].dwSize, szComma);
            break;
         }

         case TYPE_GUID:
         {
            WriteTextFile(hFmtFile, "%u SQLNCHAR 0 0 %s %u %S %s\r\n", dwColumnIndex, szDelimiter, dwColumnIndex, pOutputBase->pEntries[i].szName, STR_FMT_SQLSORT);
            WriteTextFile(hSqlFile, "   [%S] uniqueidentifier ROWGUIDCOL DEFAULT(NULL)%s\n", pOutputBase->pEntries[i].szName, szComma);
            break;
         }

         case TYPE_STRING_TO_BOOL:
         case TYPE_BOOL:
         {
            WriteTextFile(hFmtFile, "%u SQLNCHAR 0 0 %s %u %S %s\r\n", dwColumnIndex, szDelimiter, dwColumnIndex, pOutputBase->pEntries[i].szName, STR_FMT_EMPTY);
            WriteTextFile(hSqlFile, "   [%S] bit DEFAULT(NULL)%s\n", pOutputBase->pEntries[i].szName, szComma);
            break;
         }
      }

      dwColumnIndex++;
   }

   WriteTextFile(hSqlFile, ")\nGO\n\n", pOutputBase->szName);

   WriteTextFile(hSqlFile, "BULK INSERT[%S]..[%S]\n", g_GlobalConfig.szDatabaseName, pOutputBase->szName);
   WriteTextFile(hSqlFile, "   FROM '%S\\%S.tsv'\n", g_GlobalConfig.szOutputDirectory, pOutputBase->szName);
   WriteTextFile(hSqlFile, "   WITH\n   (\n");
   WriteTextFile(hSqlFile, "      FORMATFILE = '%S\\%S.fmt',\n", g_GlobalConfig.szOutputDirectory, pOutputBase->szName);
   WriteTextFile(hSqlFile, "      DATAFILETYPE = 'widechar',\n");
   WriteTextFile(hSqlFile, "      CODEPAGE = 'RAW',\n");
   WriteTextFile(hSqlFile, "      TABLOCK\n");
   WriteTextFile(hSqlFile, ")\nGO\n\n");

   return TRUE;
}

BOOL
OutputInitialize (
)
{
   for (DWORD i = 0; i < sizeof(OutputBases) / sizeof(OUTPUT_BASE); i++)
   {
      WCHAR szFilename[MAX_PATH];
      BYTE pbBomUTF16LE[4] = { 0xFF, 0xFE, 0, 0 };

      swprintf_s(szFilename, MAX_PATH, L"%s\\%s.tsv", g_GlobalConfig.szOutputDirectory, OutputBases[i].szName);
      OutputBases[i].hFile = CreateFile(szFilename, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, 0, NULL);

      if (OutputBases[i].hFile == INVALID_HANDLE_VALUE)
      {
         Log(
            0, LOG_LEVEL_ERROR,
            L"CreateFile(%s) failed (error %u).", szFilename, GetLastError()
         );
         OutputBases[i].hFile = NULL;
         return FALSE;
      }

      // Write UTF-16 BOM
      pOutputWriteFile(&OutputBases[i], (LPWSTR)pbBomUTF16LE, NULL);

      // Misc
      for (DWORD j = 0; j < OutputBases[i].dwEntriesCount; j++)
      {
         OutputBases[i].pEntries[j].dwSize = 1;             // At least one to avoid nvarchar(0)
      }
   }

   return TRUE;
}

BOOL
OutputWrite (
   _In_ LONG lGpoId,
   _In_opt_z_ LPCWSTR szType,
   _In_z_ LPCWSTR szBase,
   _In_ DWORD dwParamCount,
   ...
)
{
   va_list argptr;

   POUTPUT_BASE pOutputBase = NULL;

   for (DWORD i = 0; i < sizeof(OutputBases) / sizeof(OUTPUT_BASE); i++)
   {
      if (_wcsicmp(szBase, OutputBases[i].szName) == 0)
      {
         pOutputBase = &OutputBases[i];
         break;
      }
   }

   if (pOutputBase == NULL)
   {
      return FALSE;
   }

   if (dwParamCount != pOutputBase->dwEntriesCount)
   {
      Log(
         0, LOG_LEVEL_ERROR,
         L"Wrong number of argument."
      );
      return FALSE;
   }

   va_start(argptr, dwParamCount);

   if (pOutputBase->dwOptions & OUTPUT_OPTION_ADD_ID_GPO)
   {
      WCHAR szText[26];

      swprintf_s(szText, 26, L"%d", lGpoId);
      pOutputWriteFile(pOutputBase, szText, NULL);
      pOutputWriteFile(pOutputBase, (LPWSTR)L"\t", NULL);
   }

   if (pOutputBase->dwOptions & OUTPUT_OPTION_ADD_SCOPE)
   {
      if (szType != NULL)
         pOutputWriteFile(pOutputBase, (LPWSTR)szType, NULL);
      pOutputWriteFile(pOutputBase, (LPWSTR)L"\t", NULL);
   }

   for (DWORD i=0; i< pOutputBase->dwEntriesCount; i++)
   {
      if (i != 0)
         pOutputWriteFile(pOutputBase, (LPWSTR)L"\t", NULL);

      switch (pOutputBase->pEntries[i].Type)
      {
         case TYPE_INTEGER:
         {
            int i;
            WCHAR szText[26];

            i = va_arg(argptr, int);
            swprintf_s(szText, 26, L"%d", i);
            pOutputWriteFile(pOutputBase, szText, NULL);
            break;
         }

         case TYPE_STRING_TO_INTEGER:
         case TYPE_STRING:
         case TYPE_DATE:
         {
            LPWSTR szValue;

            szValue = va_arg(argptr, LPWSTR);

            if (szValue != NULL)
               pOutputWriteFile(pOutputBase, szValue, &pOutputBase->pEntries[i].dwSize);
            break;
         }

         case TYPE_STRING_TO_BOOL:
         {
            LPWSTR szValue;

            szValue = va_arg(argptr, LPWSTR);

            if (szValue != NULL)
            {
               if (_wcsicmp(szValue, L"true") == 0)
                  pOutputWriteFile(pOutputBase, (LPWSTR)L"1", NULL);
               else if (_wcsicmp(szValue, L"1") == 0)
                  pOutputWriteFile(pOutputBase, (LPWSTR)L"1", NULL);
               else if (_wcsicmp(szValue, L"false") == 0)
                  pOutputWriteFile(pOutputBase, (LPWSTR)L"0", NULL);
               else if (_wcsicmp(szValue, L"0") == 0)
                  pOutputWriteFile(pOutputBase, (LPWSTR)L"0", NULL);
               else           // TODO: print warning
                  pOutputWriteFile(pOutputBase, (LPWSTR)L"0", NULL);
            }
            break;
         }

         case TYPE_GUID:
         {
            HRESULT hr;
            LPGUID pGuid;
            LPWSTR szGuid;

            pGuid = va_arg(argptr, LPGUID);
            hr = StringFromCLSID(*pGuid, &szGuid);
            if (hr == S_OK)
            {
               pOutputWriteFile(pOutputBase, szGuid, NULL);
               CoTaskMemFree(szGuid);
            }
            break;
         }

         case TYPE_BOOL:
         {
            BOOL bValue;

            bValue = va_arg(argptr, BOOL);

            if (bValue == TRUE)
               pOutputWriteFile(pOutputBase, (LPWSTR)L"1", NULL);
            else
               pOutputWriteFile(pOutputBase, (LPWSTR)L"0", NULL);
         }
      }
   }

   va_end(argptr);

   pOutputWriteFile(pOutputBase, (LPWSTR)L"\r\n", NULL);

   return TRUE;
}

BOOL
OutputClose (
)
{
   WCHAR szFilename[MAX_PATH];

   HANDLE hSqlFile;

   swprintf_s(szFilename, MAX_PATH, L"%s\\import.sql", g_GlobalConfig.szOutputDirectory);

   hSqlFile = CreateFile(szFilename, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, 0, NULL);

   if (hSqlFile == INVALID_HANDLE_VALUE)
   {
      Log(
         0, LOG_LEVEL_ERROR,
         L"CreateFile(%s) failed (error %u).", szFilename, GetLastError()
      );
      return FALSE;
   }

   //
   // Generic SQL commands
   //
   WriteTextFile(hSqlFile, "USE[master]\nGO\n\n");
   WriteTextFile(hSqlFile, "DROP DATABASE IF EXISTS[%S]\nGO\n\n", g_GlobalConfig.szDatabaseName);
   WriteTextFile(hSqlFile, "CREATE DATABASE[%S]\nCOLLATE SQL_Latin1_General_CP1_CI_AS\nGO\n\n", g_GlobalConfig.szDatabaseName);
   WriteTextFile(hSqlFile, "ALTER DATABASE[%S] SET RECOVERY SIMPLE\nGO\n\n", g_GlobalConfig.szDatabaseName);
   WriteTextFile(hSqlFile, "USE[%S]\nGO\n\n", g_GlobalConfig.szDatabaseName);

   //
   // Process tables
   //
   for (DWORD i = 0; i < sizeof(OutputBases) / sizeof(OUTPUT_BASE); i++)
   {
      HANDLE hFmtFile;

      swprintf_s(szFilename, MAX_PATH, L"%s\\%s.fmt", g_GlobalConfig.szOutputDirectory, OutputBases[i].szName);

      hFmtFile = CreateFile(szFilename, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, 0, NULL);

      if (hFmtFile == INVALID_HANDLE_VALUE)
      {
         Log(
            0, LOG_LEVEL_ERROR,
            L"CreateFile(%s) failed (error %u).", szFilename, GetLastError()
         );
         break;
      }

      pCreateSqlFile(hSqlFile, hFmtFile, &OutputBases[i]);

      CloseHandle(OutputBases[i].hFile);
      CloseHandle(hFmtFile);
   }

   CloseHandle(hSqlFile);

   return TRUE;
}