#include "Convert2Access.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#ifdef _WIN32
#include <io.h>
#else
#include <unistd.h>
#endif

#include <iostream>
#include <conio.h>
#include <assert.h>

#include "libxl.h"

static char  stringSeparator = '\'';
static char *fieldSeparator = ",";
static char *lineSeparator = "\t";
static char *lineCompleteSeparator = ";";
static char *sentenceCompleteSeparator = ";";

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
CConvert2Access::CConvert2Access(CXlsx2DBTool *app)
	: _app(app)
{
	_accSql = new CAccessSql(_app);
	_accSql->Open();
}

CConvert2Access::~CConvert2Access() {
	_accSql->Close();
	delete(_accSql);
}

int
CConvert2Access::Convert(
	const std::string& sUrl,
	const std::string& sUserName,
	std::string& sPassword,
	std::string& sSource,
	config_item_t *item) {

	char inputFileName[256];
	TCHAR bookName[256] = { 0 };
	char sheetName[256] = { 0 };
	int sheetIdx;
	int titleRowId;
	int cellRow = 0, cellCol = 0;
	meta_table_t& meta_table = item->metaTable;

	libxl::Book *book;
	libxl::Sheet *sheet;
	bool bLoad;
	const char *errMsg;
	int sheetCount = 0, sheetType = 0;

	char chSqlDelete[1024] = { 0 };
	char chSqlInsert[1024 * 16] = { 0 };

	sprintf(inputFileName, "%s", item->inputName);
	sheetIdx = item->inputSheetIdx;
	titleRowId = item->inputTitleRowId;

	book = xlCreateXMLBook();
	if (!book) {
		fprintf(stderr, "libxl lib crashed!!!\n");
		return -1;
	}

#ifdef _UNICODE
	CXlsx2DBTool::C2w(inputFileName, bookName, 256);
	bLoad = book->load(bookName);
#else
	bLoad = book->load(inputFileName);
#endif
	errMsg = book->errorMessage();
	if (!bLoad || 0 != strcmp("ok", errMsg)) {
		fprintf(stderr, "file(%s) load failed, because %s!!!\n", inputFileName, errMsg);
		return -2;
	}

	// process sheet at index
	sheetCount = book->sheetCount();
	sheetType = book->sheetType(sheetIdx);
	if (sheetCount <= 0
		|| libxl::SHEETTYPE_SHEET != sheetType) {
		fprintf(stderr, "can't find sheet or sheet type error -- sheetIdx(%d), sheetCount(%d), sheetType(%d)!!!\n", sheetIdx, sheetCount, sheetType);
		return -3;
	}
	// open and parse the sheet
	sheet = book->getSheet(sheetIdx);

	// sheet name
#ifdef _UNICODE
	CXlsx2DBTool::W2c(sheet->name(), sheetName, 256);
#else
	sprintf(sheetName, sheet->name());
#endif
	
	//
	printf("\nStart dumping \"%s\"(%s):\n\n", item->inputName, sheetName);

	// table name
	if (strlen(item->sqlTableName) < 1) {
		sprintf(item->sqlTableName, "%s", sheetName);
	}

	int sheetLastRow = sheet->lastRow();
	int sheetFirstCol = meta_table.fromColumn;
	int sheetLastCol = meta_table.toColumn + 1;

	// find first all empty row as last row
	for (cellRow = 0; cellRow < sheetLastRow; ++cellRow) {
		bool bAllEmpty = true;
		// walk cells
		for (cellCol = sheetFirstCol; cellCol < sheetLastCol; ++cellCol) {
			libxl::CellType eCellType = sheet->cellType(cellRow, cellCol);
			if (libxl::CELLTYPE_BLANK != eCellType
				&& libxl::CELLTYPE_EMPTY != eCellType) {
				// check empty string
				if (libxl::CELLTYPE_STRING == eCellType) {
#ifdef _UNICODE
					const TCHAR *tstr = sheet->readStr(cellRow, cellCol);
					char buff[256] = { 0 };
					char *str = CXlsx2DBTool::W2c(tstr, buff, sizeof(buff));
#else
					const TCHAR *str = sheet->readStr(cellRow, cellCol);
#endif
					const TCHAR *sCell = sheet->readStr(cellRow, 0);
					size_t nLen = strlen(str);
					if (nLen > 0) {
						bAllEmpty = false;
						break;
					}
				}
				else {
					bAllEmpty = false;
					break;
				}
			}
		}

		if (bAllEmpty) {
			sheetLastRow = cellRow;
			break;
		}
	}

	// check data row exist or not
	if (sheetLastRow < item->inputTitleRowId + 1) {
		// table is empty
		if (strlen(item->deleteClause) >= 3) {
			// start line -- DELETE FROM xxxx
			sprintf(chSqlDelete, "DELETE FROM `%s` WHERE %s;"
				, item->sqlTableName
				, item->deleteClause);
		}

		// access delete
		if (strlen(chSqlDelete) > 1)
			_accSql->Execute(chSqlDelete);
	}
	else {
		//insert ptr
		char *iptr = chSqlInsert; // insert ptr
		size_t insert_value_pos = 0; // insert cell value position

		// at least exist one data row
		// process all rows of the sheet
		for (cellRow = 0; cellRow < sheetLastRow; ++cellRow) {
			const TCHAR *sCell = sheet->readStr(cellRow, 0);
			if (sCell && CXlsx2DBTool::StartsWith(sCell, (TCHAR *)"#")) {
				// ignore comment rows
				++meta_table.commentRowNum;
				continue;
			}

			// skip rows before title
			if (cellRow + 1 < titleRowId)
				continue;

			// walk cells
			for (cellCol = sheetFirstCol; cellCol < sheetLastCol; ++cellCol) {
				field_meta_t& meta = meta_table.arr[cellCol];

				// skip undefined field
				if (!meta.isField)
					continue;

				// collect title
				if (!meta_table.bTitleOk) {
					if (meta.isField) {
						libxl::CellType eCellType = sheet->cellType(cellRow, cellCol);
						if (libxl::CELLTYPE_STRING == eCellType) {
#ifdef _UNICODE
							const TCHAR *tstr = sheet->readStr(cellRow, cellCol);
							char buff[256] = { 0 };
							char *str = CXlsx2DBTool::W2c(tstr, buff, sizeof(buff));
#else
							const TCHAR *str = sheet->readStr(cellRow, cellCol);
#endif
							sprintf(meta.name, "%s", str);
						}
					}

					//
					if (meta.isToColumn) {
						cellCol = sheetLastCol;
						meta_table.bTitleOk = true;

						if (strlen(item->deleteClause) >= 3) {
							// start line -- DELETE FROM xxxx
							sprintf(chSqlDelete, "DELETE FROM `%s` WHERE %s;"
								, item->sqlTableName
								, item->deleteClause);
						}

						// access delete
						if (strlen(chSqlDelete) > 1)
							_accSql->Execute(chSqlDelete);

						// start line -- INSERT INTO xxx
						iptr = chSqlInsert;
						sprintf(iptr, "INSERT INTO `%s` (", item->sqlTableName);
						{
							bool bFirst = true;
							int i;
							for (i = sheetFirstCol; i < sheetLastCol; ++i) {
								field_meta_t& tmp = meta_table.arr[i];

								assert(strlen(tmp.name) > 1);

								if (tmp.isField) {
									// field name
									if (bFirst) {
										iptr = chSqlInsert + strlen(chSqlInsert);
										sprintf(iptr, "`%s`", tmp.name);
										bFirst = false;
									}
									else {
										// separate field
										iptr = chSqlInsert + strlen(chSqlInsert);
										sprintf(iptr, "%s `%s`", fieldSeparator, tmp.name);
									}
								}

								if (tmp.isToColumn)
									break;
							}
						}
						iptr = chSqlInsert + strlen(chSqlInsert);
						sprintf(iptr, ") %sVALUES ", lineSeparator);
						insert_value_pos = strlen(chSqlInsert);

					}
				}
				else {
					// output cell
					iptr = chSqlInsert + strlen(chSqlInsert);
					OutputCell(iptr, sheet, cellRow, cellCol, meta, inputFileName, sheetLastRow, sheetLastCol);
				}
			}

			// close block
			size_t szSqlLen = strlen(chSqlInsert);
			if (cellRow > 0 && szSqlLen > insert_value_pos) {
				// access instert
				_accSql->Execute(chSqlInsert);

				// carriage back
				chSqlInsert[insert_value_pos] = '\0';
			}
		}

		//
		book->release();
	}
	return 0;
}

// Output a string
void
CConvert2Access::OutputCell(char *ptr, void *sheet, int row, int col, field_meta_t& meta, const char *inputFileName, int lastRow, int lastCol) {
	char *iptr = ptr;

	libxl::Sheet *sheet_ = static_cast<libxl::Sheet *>(sheet);

	// cell type
	libxl::CellType eCellType = sheet_->cellType(row, col);

	// process none visible and empty
	if (libxl::CELLTYPE_BLANK == eCellType
		|| libxl::CELLTYPE_EMPTY == eCellType
		|| libxl::CELLTYPE_ERROR == eCellType
		|| libxl::SHEETSTATE_VISIBLE != sheet_->hidden()) {
		// output empty string
		iptr = ptr + strlen(ptr);
		OutputString(iptr, "", meta);

		//
		if (meta.isToColumn && row + 1 < lastRow) {
			iptr = ptr + strlen(ptr);
			sprintf(iptr, "%s", lineCompleteSeparator);
		}
	}
	else if (libxl::CELLTYPE_ERROR == eCellType) {
		FILE *ferr = fopen("xlsx2dbtool.error", "at+");
		fprintf(ferr, "<%s> has error cell -- row(%d)col(%d)\n", inputFileName, row, col);
		fclose(ferr);
	}
	else if (libxl::CELLTYPE_NUMBER == eCellType) {
		double number = sheet_->readNum(row, col);
		OutputNumber(iptr, number, meta);

		if (meta.isToColumn && row + 1 < lastRow) {
			iptr = ptr + strlen(ptr);
			sprintf(iptr, "%s", lineCompleteSeparator);
		}
	}
	else if (libxl::CELLTYPE_STRING == eCellType) {
#ifdef _UNICODE
		const TCHAR *tstr = sheet_->readStr(row, col);
		char buff[1024 * 16] = { 0 };
		char *str = CXlsx2DBTool::W2c(tstr, buff, sizeof(buff));
#else
		const char *str = sheet_->readStr(row, col);
#endif
		OutputString(iptr, str, meta);

		//
		if (meta.isToColumn && row + 1 < lastRow) {
			iptr = ptr + strlen(ptr);
			sprintf(iptr, "%s", lineCompleteSeparator);
		}
	}
	else if (libxl::CELLTYPE_BOOLEAN == eCellType) {
		bool b = sheet_->readBool(row, col);
		//OutputString(iptr, b ? "true" : "false", meta);
	}
	else {
		//OutputString(iptr, "", meta);
	}
}

// Output a string
void
CConvert2Access::OutputString(char *ptr, const char *str, field_meta_t& meta) {
	char *iptr = ptr;

	if (meta.isFromColumn) {
		iptr = ptr + strlen(ptr);
		sprintf(iptr, "(%c%s%c", stringSeparator, str, stringSeparator);
	}
	else {
		iptr = ptr + strlen(ptr);
		sprintf(iptr, "%s%c%s%c", fieldSeparator, stringSeparator, str, stringSeparator);
	}

	if (meta.isToColumn) {
		iptr = ptr + strlen(ptr);
		sprintf(iptr, ")");
	}
}

// Output a number
void
CConvert2Access::OutputNumber(char *ptr, const double number, field_meta_t& meta) {
	char *iptr = ptr;

	if (meta.isFromColumn) {
		sprintf(iptr, "%s(%.15g", lineSeparator, number);
	}
	else {
		sprintf(iptr, "%s%.15g", fieldSeparator, number);
	}

	if (meta.isToColumn) {
		iptr = ptr + strlen(ptr);
		sprintf(iptr, ")");
	}
}