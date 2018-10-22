#ifndef ACCESSSQL_H
#define ACCESSSQL_H

#import "C:/Program Files/Common Files/system/ado/msado15.dll" no_namespace rename("EOF", "adoEOF") rename("BOF", "adoBOF")
//#import "msado15.dll" no_namespace rename("EOF", "adoEOF") rename("BOF", "adoBOF")

#include "Xlsx2DBTool.h"

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
class CAccessSql {
public:
	CAccessSql(CXlsx2DBTool *app);
	virtual ~CAccessSql();

	BOOL Open();
	void Close();
	BOOL IsOpened();

	BOOL Execute(const char *sSql);

public:
	CXlsx2DBTool *_app;

public:
	_ConnectionPtr _pConn = nullptr;
};

#endif
