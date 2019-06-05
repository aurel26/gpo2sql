/*
gpo2sql
Copyright (c) 2019  aurel26

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/
#include <Windows.h>
#include <tchar.h>
#include "gpo2sql.h"

HANDLE g_hHeap;
HANDLE g_hLogFile = NULL;
BOOL g_bSupportsAnsi = FALSE;
GLOBAL_CONFIG g_GlobalConfig = { 0 };

#pragma comment(lib, "msxml6.lib")
#pragma comment(lib, "rpcrt4.lib")

int
wmain (
   int argc,
   wchar_t *argv[])
{
   BOOL bResult;
   BOOL bReturn = EXIT_FAILURE;
   HRESULT hr;

   WCHAR szFindPattern[MAX_PATH];

   WIN32_FIND_DATA FindFileData;
   HANDLE hFind, hStdOut;
   DWORD dwConsoleMode;

   LONG lGpoId = 1;

   g_hHeap = HeapCreate(0, 0, 0);
   if (g_hHeap == NULL)
      return EXIT_FAILURE;
   hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
   if (hr != S_OK)
      return EXIT_FAILURE;

   //
   // Initialize console
   //
   hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);

   // Set console output to 'ISO 8859-1 Latin 1; Western European (ISO)'
   SetConsoleOutputCP(28591);

   GetConsoleMode(hStdOut, &dwConsoleMode);
   g_bSupportsAnsi = SetConsoleMode(hStdOut, dwConsoleMode | ENABLE_VIRTUAL_TERMINAL_PROCESSING);
   SetConsoleTitle(L"gpo2sql");

   //
   // Read parameters
   //
   if (argc != 4)
   {
      Log(
         0, LOG_LEVEL_CRITICAL,
         L"[!] Wrong number of arguments"
      );
      Log(
         0, LOG_LEVEL_INFORMATION,
         L"[.] Usage: gpo2sql <input dir> <output dir> <database name>"
      );
      goto Fail;
   }
   else
   {
      g_GlobalConfig.szDatabaseName = argv[3];
      g_GlobalConfig.szInputDirectory = argv[1];
      g_GlobalConfig.szOutputDirectory = argv[2];
   }

   //
   // Process
   //
   bResult = OutputInitialize();
   if (bResult == FALSE)
      goto End;

   swprintf_s(szFindPattern, MAX_PATH, L"%s\\*.xml", g_GlobalConfig.szInputDirectory);

   hFind = FindFirstFile(szFindPattern, &FindFileData);
   if (hFind == INVALID_HANDLE_VALUE)
   {
      //printf("FindFirstFile failed (%d)\n", GetLastError());
   }
   else
   {
      do
      {
         if ((FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
         {
            TCHAR szFullPath[MAX_PATH];

            _stprintf_s(szFullPath, MAX_PATH, TEXT("%s\\%s"), g_GlobalConfig.szInputDirectory, FindFileData.cFileName);

            XmlParseFile(&lGpoId, szFullPath);
            lGpoId++;
         }
      } while (FindNextFile(hFind, &FindFileData) != 0);

      FindClose(hFind);
   }

   bReturn = EXIT_SUCCESS;
End:
   OutputClose();
Fail:
   CoUninitialize();
   HeapDestroy(g_hHeap);

   return bReturn;
}