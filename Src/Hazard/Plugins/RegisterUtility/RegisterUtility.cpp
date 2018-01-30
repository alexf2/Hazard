// RegisterUtility.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

const CATID CATID_FacValMonitors = 
 { 0x6e8bccf5, 0x6ae0, 0x11d4, { 0x8f, 0xba, 0x0, 0x50, 0x4e, 0x2, 0xc3, 0x9d } };

enum EN_Modes {
   ENM_Reg,
   ENM_UnReg,
   ENM_None
 };


int RU_Action( LPSTR lpPath, EN_Modes mode );
void ReportWin32Error();

void PrintHelp()
 {
   USES_OEM_XCONSOLE_CONVERSION;

   printf( W2X(L"Требуется:\n"
	   L"\tRegisterUtility.exe /key path\n"
	   L"\t\t/key - /r или /u\n"
	   L"\t\t\t/r - зарегистрировать диспетчер оценки значений ФО;\n"
	   L"\t\t\t/u - удалить диспетчер оценки значений ФО.\n"
	   L"\t\tpath - полный путь к регистрационному файлу диспетчера\n"

	   L"Пример использования:\n"
	   L"\tRegisterUtility.exe /r \"C:\\Program Files\\AlexCorp\\GasPipeline.rg\"\n\n") );
 }

int EndRoutine( int iRetCode )
 {
   USES_OEM_XCONSOLE_CONVERSION;

   printf( W2X(L"Press any key...\n") );
   getch();

   return iRetCode;
 }

LPSTR FirstNotSpace( LPSTR lp, bool bFromLeft )
 {
   if( bFromLeft )
	{
      for( ; *lp; ++lp )
	    if( !isspace(*lp) ) break;
	  return *lp ? lp:NULL;
	}
   else
	{ 
	  for( LPSTR lp2 = lp + strlen(lp) - 1; lp2 >= lp; --lp2 )
	    if( !isspace(*lp2) ) break;
	  return lp2 >= lp ? lp2:NULL;
	}
 }

void RemoveControlAndTrimL( LPSTR pT )
 {
   if( pT )
	{
	  LPSTR lpTmp = (LPSTR)_alloca( strlen(pT) + 1 );
	  LPSTR lpTmp0 = lpTmp;
	  for( LPSTR p = pT; *p; ++p )
	    if( !iscntrl(*p) ) *lpTmp++ = *p;

	   *lpTmp = 0;
	   LPSTR lpL = FirstNotSpace( lpTmp0, true );
	   LPSTR lpR = FirstNotSpace( lpTmp0, false );
	   if( !lpL ) lpL = lpTmp0;
	   if( !lpR ) lpR = lpTmp0 + strlen(lpTmp0) - 1;
	   strncpy( pT, lpL, lpR - lpL + 1), *(pT + (lpR - lpL) + 1) = 0;
	}
 }

char* FGetsSkip( char* lp, int iN, FILE* f )
 {
   char* pT;
   while( 1 )
	if( (pT = fgets(lp, iN, f)) )
	 {
	   int iSz = strlen( lp );
	   if( iSz < 1 ) continue;
       for( int i = 0; i < iSz; ++i )
		 if( !isspace(lp[ i ]) ) break;

	   if( i < iSz ) break;
	 }
	else break;

   RemoveControlAndTrimL( pT );

   return pT;
 }

int main(int argc, char* argv[])
 {
   USES_OEM_XCONSOLE_CONVERSION;
   USES_CONVERSION;
   setlocale( LC_ALL, "Russian" );

   char cTmpBuf[ 1024 + MAX_PATH ];


   TCOMInit ciInit( COINIT_APARTMENTTHREADED );
   if( !ciInit )
	{
	  _snprintf( cTmpBuf, sizeof(cTmpBuf), "Ошибка инициализации COM: HRESULT = 0x%x\n", (int)(HRESULT)ciInit );
	  printf( W2X(A2W(cTmpBuf)) );

	  return EndRoutine( -1 );
	}
      

   if( argc != 3 )
	{
	  printf( W2X(L"Ошибочные параметры. ") );
	  PrintHelp();	  

	  return EndRoutine( -1 );
	}
   

   EN_Modes enMode;
   if( !_stricmp("/r", argv[1]) ) enMode = ENM_Reg;
   else if( !_stricmp("/u", argv[1]) ) enMode = ENM_UnReg;
   else 
	{
      _snprintf( cTmpBuf, sizeof(cTmpBuf), "Ошибочный ключ: %s. ", argv[1] );
	  printf( W2X(A2W(cTmpBuf)) );
	  PrintHelp();

	  return EndRoutine( -1 );
	}

   FHolder f;
   f = fopen( argv[2], "r" );
   if( !f )
	{
	  _snprintf( cTmpBuf, sizeof(cTmpBuf), "Ошибка открытия файла: [%s]\n", argv[2] );
	  printf( W2X(A2W(cTmpBuf)) );
	  _snprintf( cTmpBuf, sizeof(cTmpBuf), "\t%s\n", _strerror(NULL) );
	  printf( W2X(A2W(cTmpBuf)) );

	  return EndRoutine( -1 );
	}

   char chPathDLL0[ MAX_PATH ], chPathDLL[ MAX_PATH ];
   if( !FGetsSkip(chPathDLL0, sizeof(chPathDLL0) - 1, f) )
	{	  
	  _snprintf( cTmpBuf, sizeof(cTmpBuf), "Неожиданный конец или ошибка чтения файла: [%s]\n", argv[2] );
	  printf( W2X(A2W(cTmpBuf)) );
	  _snprintf( cTmpBuf, sizeof(cTmpBuf), "\t%s\n", _strerror(NULL) );
	  printf( W2X(A2W(cTmpBuf)) );

	  return EndRoutine( -1 );
	}

   int iResA;
   if( enMode == ENM_Reg )
	{
	  iResA = RU_Action( chPathDLL0, enMode );
	  if( iResA ) return EndRoutine( iResA );
	}

   CComPtr<ICatRegister> spCat;
   HRESULT hr = spCat.CoCreateInstance( CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER );
   
   
   if( FAILED(hr) )
	{
	  printf( W2X(L"\nОшибка создания менеджера категорий компонент\n") );
	  return EndRoutine( -1 );
	}

   while( 1 )
	{
	  GUID guTmp;
      if( FGetsSkip(chPathDLL, sizeof(chPathDLL) - 1, f) )
	   {	     
		 hr = CLSIDFromProgID( A2W(chPathDLL), &guTmp );
		 if( FAILED(hr) )
		  {
		    _snprintf( cTmpBuf, sizeof(cTmpBuf), "\tОшибка извлечения CLSID для: %s\n", chPathDLL );
			printf( W2X(A2W(cTmpBuf)) );
		    continue;
		  }
		 else
		  {		    			  
            if( enMode == ENM_Reg )
			 {
			   hr = spCat->RegisterClassImplCategories( guTmp, 1, &(CATID)CATID_FacValMonitors );
			   _snprintf( cTmpBuf, sizeof(cTmpBuf), "\tНазначение категории для: %s - %s\n", chPathDLL, SUCCEEDED(hr) ? "успешно":"ошибка" );
			 }
			else
			 {
			   hr = spCat->UnRegisterClassImplCategories( guTmp, 1, &(CATID)CATID_FacValMonitors );
			   _snprintf( cTmpBuf, sizeof(cTmpBuf), "\tУдаление категории для: %s - %s\n", chPathDLL, SUCCEEDED(hr) ? "успешно":"ошибка" );
			 }			
			printf( W2X(A2W(cTmpBuf)) );
		  }
	   }
	  else
	   {
	     if( feof((FILE*)f) ) break;
		 else
		  {
		    _snprintf( cTmpBuf, sizeof(cTmpBuf), "Ошибка чтения файла: [%s]\n", argv[2] );
			printf( W2X(A2W(cTmpBuf)) );
			_snprintf( cTmpBuf, sizeof(cTmpBuf), "\t%s\n", _strerror(NULL) );
			printf( W2X(A2W(cTmpBuf)) );

			return EndRoutine( -1 );
		  }
	   }
	}

   if( enMode == ENM_UnReg )
	{
	  iResA = RU_Action( chPathDLL0, enMode );
	  if( iResA ) return EndRoutine( iResA );
	}

   return EndRoutine( 0 );
 }

int RU_Action( LPSTR lpPath, EN_Modes mode )
 {
   USES_OEM_XCONSOLE_CONVERSION;
   USES_CONVERSION;

   char cTmpBuf[ 1024 + MAX_PATH ];

   HINSTANCE hLib = LoadLibrary( lpPath );

   if( hLib < (HINSTANCE)HINSTANCE_ERROR )
    {
	  ReportWin32Error();
	  return -1;
    }

   typedef HRESULT  (WINAPI *TRegProc)();
   TRegProc lpDllEntryPoint;
   HRESULT hr;

   lpDllEntryPoint = (TRegProc)GetProcAddress( hLib, mode == ENM_Reg ? "DllRegisterServer":"DllUnregisterServer" );
   if( lpDllEntryPoint != NULL )
	 hr = lpDllEntryPoint();
   else 
	{
	  ReportWin32Error();
	  FreeLibrary( hLib );
	  return -1;
	}

   FreeLibrary( hLib );
   
   if( mode == ENM_Reg )
     _snprintf( cTmpBuf, sizeof(cTmpBuf), SUCCEEDED(hr) ? "\nРегистрация COM-сервера [%s] прошла успешно: [0x%x]\n":"Сбой при регистрации COM-сервера: [0x%x]\n", lpPath, (int)hr );
   else
	 _snprintf( cTmpBuf, sizeof(cTmpBuf), SUCCEEDED(hr) ? "\nОтмена регистрации COM-сервера [%s] прошла успешно: [0x%x]\n":"Сбой при регистрации COM-сервера: [0x%x]\n", lpPath, (int)hr );
   printf( W2X(A2W(cTmpBuf)) );

   return 0;
 }

void ReportWin32Error()
 {
   USES_OEM_XCONSOLE_CONVERSION;
   USES_CONVERSION;

   LPVOID lpMsgBuf = NULL;
   DWORD dwErr;
   DWORD dw = FormatMessage( 
	 FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
	 NULL,
	 (dwErr = GetLastError()),
	 MAKELANGID(LANG_RUSSIAN, SUBLANG_DEFAULT), // Default language
	 (LPTSTR) &lpMsgBuf,
	 0,
	 NULL 
   );

   if( !dw )
	 dw = FormatMessage( 
	 FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
	 NULL,
	 dwErr,
	 MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
	 (LPTSTR) &lpMsgBuf,
	 0,
	 NULL 
   );

   UINT uiSz = LocalSize( (HLOCAL)lpMsgBuf );
   LPSTR lpTmp = (LPSTR)_alloca( dw ? uiSz + 1024:1024 );
   
   sprintf( lpTmp, "%s\n", dw ? lpMsgBuf:"Невозможно извлечь информацию об ошибке" );
   printf( W2X(A2W(lpTmp)) );
   
   if( lpMsgBuf ) LocalFree( lpMsgBuf );
 }

