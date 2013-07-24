rtdfile
=======

MS Excel Real Time Data (RTD) with live TSV File sources

The RTDFile MS Excell add-in lets you automatically retrieve real-time data as it changes, from tab-delimited text files into Microsoft Excel 2002. A cell within the tab-delimited text file is retrieved via the RealTimeData (RTD) function, supplying the RTDFile add-in, the name of the tab-delimited text file, and a cell reference.
Setup

To install the RTDFile add-in, make sure Excel is closed, then run the following command.
regsvr32.exe c:\RTDFile\Debug\rtdfile.dll
Syntax

=RTD("rtdfile",,"meta")
=RTD("rtdfile",,"filename","cellref")
"rtdfile"  The quoted program name of the RTDFile add-in. Specify it exactly as shown.
"meta"	One of the following RTDFile keywords (must be quoted):
"=version"	Returns the RTDFile version
"=web"	Web address for RTDFile website.
"=help"	Web address for RTDFile help documentation.
"filename"	The quoted name of the tab-delimited text file.
"cellref"	The quoted cell reference in filename.
For additional information, see How to set up and use the RTD function

Examples

Consider the demo file stored in \RTDFile\demo\days_in_month.txt, containing the following tab-delimited lines.
January   JAN       31
February  FEB       28
March     MAR       31
April     APR       30
May       MAY       31
June      JUN       30
July      JUL       31
August    AUG       31
September SEP       30
October   OCT       31
November  NOV       30
December  DEC       31
