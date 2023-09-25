
## .xlsx file creation from Harbour 100% free of  dependencies ( dll or closed libs ) <br> All from the sources! **ONLY ONE LIB IS NEEDED!**
### Sources from:

The library is based on the original C library:
[libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter)
Code included is version 1.1.5 December 30 2022 <br>


The library was originally developed by 
[hbxlsxwriter](https://github.com/riztan/hbxlsxwriter)

then cloned (not forked?)

[hbxlsxwriter](https://github.com/diegofazio/hbxlsxwriter)


This is the private fork of Francesco based on the last repository<br>

***
### Note for this fork:
Done a lot of changes to improve the use of the library.

<b>lxw_workbook</b> and <b>lxw_format</b> structs are now proper Harbour and are correctly deleted when variable goes out of scope. They are also checked at run parameters.

Implemented a check if a parameter is not correct. Just in a few places.

There are some more missing checks, almost all parameters are not checked, so there may be errors in Harbour code that goes undetected.

<b>lxw_worksheet</b> are not checked and may break the code.

Implemented row_col_options passed as hash.


REMOVED: Binaries built in msvc(2019)/mingw(10.3)/bcc64(7.1) are included in ./lib for testing.<br> 
Tests Samples -> /test<br>
New minizip version is used. 
