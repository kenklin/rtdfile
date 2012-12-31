#include "RTDFileData.h"
#include "RTDFileDLL.h"			// RTDFile_Version
#include "time.h"
#include "Tuple.h"
#include "stdio.h"


#define EMPTY	""

static std::string Now()
{
	char		now[200];
	time_t		t;
	struct tm*	tmp;

	t = time(NULL);
	tmp = localtime(&t);
	if (tmp != NULL)
		strftime(now, sizeof(now), "%H:%M:%S", tmp);

	return now;
}



bool SplitCell(const char* Cell, int* row, int* col)
{
	bool	okay = false;
	char	c = '\0';

	if (sscanf(Cell, "%c%d", &c, row) == 2) {
		*row = *row - 1;	// 0-relative
		if (c >= 'a' && c <= 'z') {
			*col = c - 'a';
			okay = true;
		} else if (c >= 'A' && c <= 'Z') {
			*col = c - 'A';
			okay = true;
		}
	}

	return okay;
}

/**
*	Returns the value in the specified RTD tab-separated file
*
*	@param[in]	Args	A 2-tuple of filename and cell
*
*	@return		The value of the cell in the filename.  If the filename/cell
*				cannot be found, the EMPTY string is returned.
*/
std::string RTDFileData::LookupData(const std::string& Args)
{
	// See http://www.cplusplus.com/reference/string/string/
	int			argc = 0;
	std::string	out, filename, cell;

	if ((argc = Tuple::Split(Args, &filename, &cell)) > 0) {
		if (filename == "=version") {
			out = RTDFile_Version;

		} else if (filename == "=email") {
			out = "rtdfile@p1software.com";

		} else if (filename == "=web") {
			out = "http://p1sofware.com/rtdfile/";

		} else if (filename == "=help") {
			out = "http://p1sofware.com/rtdfile/help.html";

		} else if (filename == "=now") {
			out = Now();

		} else {
//			out = m_Data[Args];
//			if (out == "") {
				FILE *f = fopen(filename.c_str(), "r");
				if (f == NULL) {
					out = "Err opening '" + filename + "'";
				} else {
					if (cell == "=rows") {
						char line[256] = "";
						int rows = 0;
						while (fgets(line, sizeof(line), f) != 0) {
							rows++;
						}
						char str[20] = "";
						itoa(rows, str, 10);
						out = str;

					} else {
						int cellcol = 0;
						int cellrow = 0;
						if (SplitCell(cell.c_str(), &cellrow, &cellcol)) {
							char line[256] = "";
							for (int row = 0; row <= cellrow; row++) {
								if (fgets(line, sizeof(line), f) == 0) {
#if 0
// Don't want this to display in the Excel cell
									out = "No row ";	// No such row
									char num[10] = "";
									itoa(cellrow, num, 10);
									out += num;
#else
									out = EMPTY;
#endif
								} else if (row == cellrow) {
									int len = strlen(line);
									while (len > 0 && (line[len-1] == '\n' || line[len-1] == '\r')) {
										line[len-1] = '\0';
										len--;
									}
									if (!Tuple::Get(line, cellcol, &out)) {
#if 0
// Don't want this to display in the Excel cell
										out = "No col ";	// No such col
										char num[10] = "";
										itoa(cellcol, num, 10);
										out += num;
#else
										out = EMPTY;
#endif
									}
								}
							}
						} else {
							out = "Err reading cell '" + cell + "' in file '" + filename + "'";
						}
					}

					fclose(f);
				}
//			}
		}
	}

	return out;
}
