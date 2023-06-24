SET HB_WITH_XLSXWRITER=./libxlsxwriter
@del .\lib\x64\hbxlsxwriter.lib
@del .\lib\x32\hbxlsxwriter.lib
@hbmk2 hbxlsxwriter.hbp -rebuild
@copy .\lib\x64\hbxlsxwriter.lib \harbour\lib\win\msvc64
@copy .\lib\x32\hbxlsxwriter.lib \harbour\lib\win\msvc