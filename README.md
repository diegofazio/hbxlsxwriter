
## .xlsx file creation from Harbour 100% free of  dependencies ( dll or closed libs ) <br> All from the sources! **ONLY ONE LIB IS NEEDED!**
### Sources from:

The library is based on the original C library:
[libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter)
Code included is version 1.1.5 December 30 2022 <br>


The library was originally developed by 
[https://github.com/riztan/hbxlsxwriter](https://github.com/riztan/hbxlsxwriter)

then cloned (not forked?)

[https://github.com/diegofazio/hbxlsxwriter](https://github.com/diegofazio/hbxlsxwriter)


This is the private fork of Francesco based on this last repository<br>

***
### Note for this fork:
Done a lot of changes to improve the use of the library.

<b>lxw_workbook</b> and <b>lxw_format</b> structs are now proper Harbour variables and are correctly deleted when variable goes out of scope. They are also checked at runtime when used as parameters, so that the the C library doesn't core dumps. It reports a Harbour error.

<b>lxw_worksheet</b> are not checked and may core dump, but I actually don't see how to solve the problem. 

There are some more missing checks, almost all parameters are not checked, so there may be errors in Harbour code that goes undetected. Probably the library should raise an Harbour error if the C library reports an error.


Implemented row_col_options passed as hash.  Some more hashes should be implemented.


REMOVED: Binaries built in msvc(2019)/mingw(10.3)/bcc64(7.1) are included in ./lib for testing.<br> 
Tests Samples -> /test<br>
New minizip version is used. 
