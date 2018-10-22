#include "AccessSql.h"

#include <stdint.h>

#define YEAR(y) ( (y) + 1900)

double
ToDouble(_variant_t & vt)
{
	double dVal = 0;
	if (vt.vt == VT_NULL)
		dVal = 0;
	else
		if (vt.vt == VT_R8)
			dVal = vt.dblVal;
		else
			dVal = (int)vt.lVal;

	return dVal;
}

int64_t
ToBigint(_variant_t & vt)
{
	int64_t nVal = 0;
	if (vt.vt == VT_NULL)
		nVal = 0;
	else
		if (vt.vt == VT_R4 || vt.vt == VT_R8)
			nVal = (int64_t)vt.lVal;
		else
			nVal = vt.lVal;
	return nVal;
}

int
ToInt(_variant_t & vt)
{
	int nVal = 0;
	if (vt.vt == VT_NULL)
		nVal = 0;
	else
		if (vt.vt == VT_R4 || vt.vt == VT_R8)
			nVal = (int)vt.lVal;
		else
			nVal = vt.lVal;
	return nVal;
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
CAccessSql::CAccessSql(CXlsx2DBTool *app)
	: _app(app)
{
	//
	CoInitialize(nullptr);
}

CAccessSql::~CAccessSql()
{
	Close();
	CoUninitialize();
}

BOOL
CAccessSql::Open()
{
	BOOL bResult = TRUE;
	HRESULT hr = S_OK;
	try
	{
		hr = _pConn.CreateInstance(_uuidof(Connection));
		if (SUCCEEDED(hr))
		{
			// Connect to Access database through JET (using Data Source Name, or DSN, and individual arguments instead of a connection string) 
			char chConn[256];
			sprintf(chConn, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=%s",
				_app->_config._sDataSource.c_str());

			wchar_t tchConn[1024];
			wchar_t tchUserName[256];
			wchar_t tchPassword[256];
			CXlsx2DBTool::C2w(chConn, tchConn, sizeof(tchConn));
			CXlsx2DBTool::C2w(_app->_config._sUserName.c_str(), tchUserName, sizeof(tchUserName));
			CXlsx2DBTool::C2w(_app->_config._sPassword.c_str(), tchPassword, sizeof(tchPassword));

			hr = _pConn->Open(tchConn, _app->_config._sUserName.c_str(), _app->_config._sPassword.c_str(), adModeUnknown);
			_pConn->CursorLocation = adUseClient;
		}
		else {
			bResult = FALSE;
		}
	}
	catch (_com_error e)
	{
#ifdef _UNICODE
		char buffErrmsg[4096] = { 0 };
		const TCHAR *terrmsg = e.ErrorMessage();
		CXlsx2DBTool::UnicodeToGB2312(buffErrmsg, sizeof(buffErrmsg), terrmsg, (int)wcslen(terrmsg));
		char *errmsg = buffErrmsg;

		char buffErrdesc[4096] = { 0 };
		const TCHAR *terrdesc = e.Description();
		CXlsx2DBTool::UnicodeToGB2312(buffErrdesc, sizeof(buffErrdesc), terrdesc, (int)wcslen(terrdesc));
		char *errdesc = buffErrdesc;
#else
		const char *errmsg = e.ErrorMessage();
		const char *errdesc = e.Description();
#endif

		printf("CAccessSql::Open(): failed(0x%08x) -- %s, %s\n", e.Error(), errmsg, errdesc);
		bResult = FALSE;

		//
		system("pause");
		exit(-1);
	}
	return bResult;
}

void
CAccessSql::Close()
{
	if (IsOpened())
		_pConn->Close();
	if (_pConn.GetInterfacePtr())
		_pConn.Release();
}

BOOL 
CAccessSql::IsOpened()
{
	if (!_pConn.GetInterfacePtr())
		return FALSE;
	return (_pConn->GetState() == 1);
}

BOOL 
CAccessSql::Execute(const char *sSql)
{
	BOOL bResult = TRUE;
	_RecordsetPtr rs = NULL;
	HRESULT hr = S_OK;

	wchar_t tchSql[4096];
	CXlsx2DBTool::C2w(sSql,tchSql, sizeof(tchSql));

	try
	{
		//hr = rs.CreateInstance(_uuidof(_Recordset));
		hr = rs.CreateInstance(_T("ADODB.Recordset"));
		if (SUCCEEDED(hr))
		{
			hr = rs->Open(_variant_t(tchSql),
				_variant_t((IDispatch*)_pConn, true),
				adOpenForwardOnly, adLockOptimistic, adCmdText);
			if (SUCCEEDED(hr))
			{
#ifdef _UNICODE
				char buff[4096] = { 0 };
				CXlsx2DBTool::UnicodeToGB2312(buff, sizeof(buff), tchSql, (int)wcslen(tchSql));
				char *sqlmsg = buff;
#else
				char *sqlmsg = sSql;
#endif

				// Do something here ...
				printf("Success: %s\n", sqlmsg);
			}
		}
		rs = NULL;
	}
	catch (_com_error e)
	{
#ifdef _UNICODE
		char buffErrmsg[4096] = { 0 };
		const TCHAR *terrmsg = e.ErrorMessage();
		CXlsx2DBTool::UnicodeToGB2312(buffErrmsg, sizeof(buffErrmsg), terrmsg, (int)wcslen(terrmsg));
		char *errmsg = buffErrmsg;

		char buffErrdesc[4096] = { 0 };
		const TCHAR *terrdesc = e.Description();
		CXlsx2DBTool::UnicodeToGB2312(buffErrdesc, sizeof(buffErrdesc), terrdesc, (int)wcslen(terrdesc));
		char *errdesc = buffErrdesc;
#else
		const char *errmsg = e.ErrorMessage();
		const char *errdesc = e.Description();
#endif

		printf("CAccessSql::Execute(): failed(0x%08x) -- %s, %s\n", e.Error(), errmsg, errdesc);
		bResult = FALSE;

		//
		system("pause");
		exit(-1);
	}
	return bResult;
}