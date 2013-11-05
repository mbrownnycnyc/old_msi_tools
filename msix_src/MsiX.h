
#pragma once


#define WIN32_LEAN_AND_MEAN

#include <tchar.h>
#include <stdio.h>
#include <stdlib.h>
#include <crtdbg.h>
#include <windows.h>
#include <objbase.h>
#include <msiquery.h>
#include <exception>
#include <fstream>

#include <codeanalysis/sourceannotations.h>
using namespace vc_attributes;

#define SAFE_RELEASE(ptr) if (ptr) { (ptr)->Release(); ptr = NULL; }
#define COLOR_ERROR FOREGROUND_RED | FOREGROUND_INTENSITY

void usage([Pre(Null=No)] LPCTSTR pszPath, [Pre(Null=No)] FILE* out);
void error([Pre(Null=No)] LPCTSTR pszFormat, ...);
void win32_error(DWORD dwError);

typedef struct _ARGS
{
    LPTSTR Path;
    LPTSTR Directory;
    BOOL   IncludeExtension;
} ARGS, *LPARGS;

DWORD ParseArguments(int argc, [Pre(Null=No)] _TCHAR* argv[], [Pre(Null=No)] LPARGS args);
LPTSTR MakePath([Pre(Null=No)] LPTSTR pszDest, size_t cchDest, [Pre(Null=Maybe)] LPCTSTR pszDir,
    [Pre(Null=No)] LPCTSTR pszName, [Pre(Null=Maybe)] LPCTSTR pszExt);
UINT GetString(MSIHANDLE hRecord, UINT iField,
    [Pre(Null=No)] LPTSTR* ppszProperty, [Pre(Null=No)] DWORD* pcchProperty);
HRESULT SaveStorage([Pre(Null=No)] IStorage* pRootStorage, [Pre(Null=Maybe)] LPCTSTR pszDir,
    [Pre(Null=No)] PCWSTR pszName, [Pre(Null=Maybe)] LPCTSTR pszExt);
UINT SaveStream(MSIHANDLE hRecord, [Pre(Null=Maybe)] LPCTSTR pszDir, BOOL fIncludeExt);
LPTSTR MakePathForData([Pre(Null=No)] LPTSTR pszDest, size_t cchDest,
    [Pre(Null=Maybe)] LPCTSTR pszDir, [Pre(Null=No)] LPCTSTR pszName,
    [Pre(Null=No)] LPCVOID pBuffer, size_t cbBuffer, BOOL fIncludeExt);

BOOL IsPatch([Pre(Null=No)] IStorage* pStorage);
EXTERN_C const CLSID CLSID_MsiPatch = {0xC1086, 0x0, 0x0, {0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46}};

// Simple string conversion classes.
class CW2W
{
public:

	CW2W(LPCWSTR psz) throw(...) :
		m_psz(psz)
	{
		_ASSERTE(psz);
	}

	CW2W(LPCWSTR psz, UINT nCodePage) throw(...) :
		m_psz(psz)
	{
		_ASSERTE(psz);
		(void)nCodePage;
	}

	~CW2W() throw()
	{
	}

	operator LPWSTR() const throw()
	{
		return const_cast<LPWSTR>(m_psz);
	}

private:

	LPCWSTR m_psz;

	// Hide the copy constructor and assignment operator.
	CW2W( const CW2W& ) throw();
	CW2W& operator=( const CW2W& ) throw();
};

class CW2A
{
public:

	CW2A(LPCWSTR psz) throw(...) :
		m_psz(NULL)
	{
		_ASSERTE(psz);
		Init(psz, CP_ACP);
	}

	CW2A(LPCWSTR psz, UINT nCodePage) throw(...) :
		m_psz(NULL)
	{
		_ASSERTE(psz);
		Init(psz, nCodePage);
	}

	~CW2A() throw()
	{
		if (m_psz)
		{
			delete [] m_psz;
			m_psz = NULL;
		}
	}

	operator LPSTR() const throw()
	{
		return m_psz;
	}

private:

	LPSTR m_psz;

	void Init(LPCWSTR psz, UINT nCodePage) throw(...)
	{
		int iErr = NOERROR;
		int cch = 0;

		cch = WideCharToMultiByte(nCodePage, 0, psz, -1, NULL, 0, NULL, NULL);
		if (cch)
		{
			// cch includes the NULL terminator character.
			m_psz = new CHAR[cch]; // Throws if no memory.
			if (!WideCharToMultiByte(nCodePage, 0, psz, -1, m_psz, cch, NULL, NULL))
            {
                iErr = (int)GetLastError();
            }
		}
		else
		{
			iErr = (int)GetLastError();
		}

		if (ERROR_SUCCESS != iErr)
		{
			throw iErr;
		}
	}

	// Hide the copy constructor and assignment operator.
	CW2A( const CW2A& ) throw();
	CW2A& operator=( const CW2A& ) throw();
};

class CA2W
{
public:

	CA2W(LPCSTR psz) throw(...) :
		m_psz(NULL)
	{
		_ASSERTE(psz);
		Init(psz, CP_ACP);
	}

	CA2W(LPCSTR psz, UINT nCodePage) throw(...) :
		m_psz(NULL)
	{
		_ASSERTE(psz);
		Init(psz, nCodePage);
	}

	~CA2W() throw()
	{
		if (m_psz)
		{
			delete [] m_psz;
			m_psz = NULL;
		}
	}

	operator LPWSTR() const throw()
	{
		return m_psz;
	}

private:

	LPWSTR m_psz;

	void Init(LPCSTR psz, UINT nCodePage) throw(...)
	{
		int iErr = NOERROR;
		int cch = 0;

		cch = MultiByteToWideChar(nCodePage, MB_PRECOMPOSED, psz, -1, NULL, 0);
		if (cch)
		{
			// cch includes the NULL terminator character.
			m_psz = new WCHAR[cch]; // Throws if no memory.
			if (!MultiByteToWideChar(nCodePage, MB_PRECOMPOSED, psz, -1, m_psz, cch))
            {
                iErr = (int)GetLastError();
            }
		}
		else
		{
			iErr = (int)GetLastError();
		}

		if (ERROR_SUCCESS != iErr)
		{
			throw iErr;
		}
	}

	// Hide the copy constructor and assignment operator.
	CA2W( const CA2W& ) throw();
	CA2W& operator=( const CA2W& ) throw();
};

class CA2A
{
public:

	CA2A(LPCSTR psz) throw(...) :
		m_psz(psz)
	{
		_ASSERTE(psz);
	}

	CA2A(LPCSTR psz, UINT nCodePage) throw(...) :
		m_psz(psz)
	{
		_ASSERTE(psz);
		(void)nCodePage;
	}

	~CA2A() throw()
	{
	}

	operator LPSTR() const throw()
	{
		return const_cast<LPSTR>(m_psz);
	}

private:

	LPCSTR m_psz;

	// Hide the copy constructor and assignment operator.
	CA2A( const CA2A& ) throw();
	CA2A& operator=( const CA2A& ) throw();
};

// Declare macros for proper conversion
// based on whether UNICODE being defined.
#ifdef UNICODE
#define CT2W CW2W
#define CT2A CW2A
#define CW2T CW2W
#define CA2T CA2W
#else
#define CT2W CA2W
#define CT2A CA2A
#define CW2T CW2A
#define CA2T CA2A
#endif // UNICODE
