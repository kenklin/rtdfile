rtdfile
=======

MS Excel Real Time Data (RTD) with live TSV File sources

The RTDFile MS Excel add-in lets you automatically retrieve real-time data as it changes, from 
[tab-delimited text](http://office.microsoft.com/en-us/excel-help/file-formats-that-are-supported-in-excel-HP010014103.aspx#TextFormats)
files into Microsoft Excel 2002. A cell within the tab-delimited text file is retrieved via the 
[RealTimeData](http://office.microsoft.com/en-us/excel-help/rtd-HP003066237.aspx) (RTD) function,
supplying the RTDFile add-in, the name of the tab-delimited text file, and a cell reference.

Setup
-----
To install the RTDFile add-in, make sure Excel is closed, then run the following command.

    regsvr32.exe c:\RTDFile\Debug\rtdfile.dll

Syntax
------
`=RTD("rtdfile",,` *meta* `)`

`=RTD("rtdfile",,` *filename* `,` *cellref* `)`

- "rtdfile"  The quoted program name of the RTDFile add-in. Specify it exactly as shown.
- *meta*	One of the following RTDFile keywords (must be quoted):
  - `"=version"`	Returns the RTDFile version
  - `"=web"`	Web address for RTDFile website.
  - `"=help"`	Web address for RTDFile help documentation.
- *filename*	The quoted name of the tab-delimited text file.
- *cellref*	The quoted cell reference in filename.

For additional information, see [How to set up and use the RTD function in Excel]
(http://support.microsoft.com/kb/289150).

Example: Old Faithful
---------------------
Examine the sample [oldfaithful.txt](https://github.com/kenklin/rtdfile/blob/master/demo/oldfaithful.txt) data file.
```
    Row 2: Date        Duration        Interruption Time
    Row 3: 1        4.4        78
    Row 4: 1        3.9        74
```
The corresponding cells in the [oldfaithful.xls](https://github.com/kenklin/rtdfile/blob/master/demo/oldfaithful.xls)
Excel file has these contents, whose computes values change as the oldfaithful.txt file changes.  The Excel file
uses its charting feature to graph the data.
```
    A3: Day
    B3: Duration (mins)
    C3: Interruption (mins)

    A4: =RTD("rtdfile",,file,ADDRESS(ROW()-1,COLUMN(),4))</td>
    B4: =VALUE(RTD("rtdfile",,file,ADDRESS(ROW()-1,COLUMN(),4)))</td>
    C4: =VALUE(RTD("rtdfile",,file,ADDRESS(ROW()-1,COLUMN(),4)))</td>
    D4: =ROW()-3</td>
    
    A5: =RTD("rtdfile",,file,ADDRESS(ROW()-1,COLUMN(),4))</td>
    B5: =VALUE(RTD("rtdfile",,file,ADDRESS(ROW()-1,COLUMN(),4)))</td>
    C5: =VALUE(RTD("rtdfile",,file,ADDRESS(ROW()-1,COLUMN(),4)))</td>
    D5: =ROW()-3</td>
````
![Screenshot](https://raw.github.com/kenklin/rtdfile/master/demo/oldfaithful.png)

Backstory
---------
In 2010-2012 [Grand-Am](http://grand-am.com) race seasons, I worked weekends with 
[Phoenix Performance](http://phoenixperformance-news.com/) for the Subaru Road Racing Team.
I wrote custom software that normalized the live track information transmitted to teams.
A program parsed the [putty.log](http://putty.org) stream attached to our TransNET RS232 receiver and produced
CSV files that [RTDFile](https://github.com/kenklin/rtdfile) used to populated our Excel spreadsheet.
A series of web-based dashboards written by 
[Michael](https://github.com/michaelelin) and 
[David](https://github.com/crazycow013)
then used this data to derive lap, time, and fuel information
to help with race strategy.  These were fun, but very tense weekends!

![](http://p1software.com/wp-content/uploads/2013/08/ken_srrt_2_small.jpg)

At the beginning of 2013, the factory [terminated funding](http://subiesport.tv/site/2013/02/subaru-terminating-road-racing-team/).
[Phoenix Performance](http://phoenixperformance-news.com/) continues to build, race, and sell championship winning cars.

In 2014 I shifted my full time efforts to [P1Software](http://p1software.com/about/) - Software, DataViz, and DataScience for Motorsports.
