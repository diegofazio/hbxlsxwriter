## .xlsx file creation from Harbour 100% free of  dependencies ( dll or closed libs ) <br> All from the sources!
### Sources

[libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter)  <-- downloaded library( 1.1.5 December 30 2022 ) - Note:minor changes<br>
[hbxlsxwriter](https://github.com/riztan/hbxlsxwriter)  <-- **riztan** libxlsxwriter Wrapper. Tests Samples -> /test

Note: Binaries for msvc64/32 and bcc64 are included for windows for testing. New minizip version for harbour is needed. Located in Lib folder ( libxlswriter/third_party/minizip )
In case you need to build all from sources, there is a build script in each folder. <br>
1) hbxlsxwriter/libxlswriter/third_party/minizip<br>
2) hbxlsxwriter/libxlswriter<br>
3) hbxlsxwriter<br>

minizip -> libxlswriter/third_party/minizip( replace harbour/lib with the one created )<br>
Test it with Visual Studio 2019( msvc64 )/bcc64 7.1
