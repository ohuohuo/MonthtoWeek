##Month to Week

This java program can transform a month-view table(including all its data) in .xlsx file into week-view table in a new .xlsx file.

The how_to_use.txt file introduce how to use it.

Wrote this code for my girlfriend so that she doesn't have to do this manually.

Rethink: 1. Actually this might be a good program example to code using multi-thread. 
	 2. This program shows a very low performance when reads a 23.2mb .xlsx file. It consumes up to 2.92GB mem and makes CPU rate to up to 95+%.
	 And finlly throws OutOfMemoryError: Java heap space exception. This problem may be soloved by changing from usermodel API to event API in Apache poi.
