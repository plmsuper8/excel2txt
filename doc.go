// test_excel project doc.go

/*
//
This CLT program transfer excel to plain txt.
	Support .xlsx on all platform, and .xls on windows (maybe buggy).
	So far, only extract Sheet1.
	working list:
		2020/11/5	UPD: print all sheets instead of first one
		2020/11/5	UPD: drop xls support, cuz old package failed to build.
		Done!.support both XLS and XLSX
		Done!.support flag (see -help)
		Done!.fix utf8 on windows by adding BOM head (-bom)
		Done!.Check xls binary head <D0 CF 11 E0 A1 B1 1A E1>
		.Multi-threads, only print finished file
		.flag sheet id, cols, rows etc.

//
Usage of test_excel.exe:
  -bom
        add byte sequence <EF BB BF> in head
                of utf8 file. Required by Microsoft, but not for Linux
  -dirname string
        the target directory or xlsx file (default "./")
  -sep string
        seperator of output (default "\t")

*/
package main
