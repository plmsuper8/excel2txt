# excel2txt
A golang command-line tool to transfer excel to plain txt


### About:

Support .xlsx on all platform, and .xls on windows (maybe buggy).

So far, only extract Sheet1
  
    

### Usage:

>> example: test_excel -bom /path/to/dir |grep -v "^#" > output.xls

>> example: test_excel /path/to/dir |grep -v "^#" > output.txt


  -bom

        add byte sequence <EF BB BF> in head
      
        of utf8 file. Required by Microsoft, but not for Linux
      
  -dirname string

        the target directory or xlsx file (default "./")
      
  -sep string

        seperator of output (default "\t")
        
### working list


> Multi-thread

> flag sheet id, cols, rows etc.
