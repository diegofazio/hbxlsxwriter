SET HB_WITH_XLSXWRITER=./libxlsxwriter
@del .\lib\x64\hbxlsxwriter.a
@del .\lib\x32\hbxlsxwriter.a
@hbmk2 hbxlsxwriter.hbp -rebuild
@copy .\lib\x64\hbxlsxwriter.a \harbour\lib\win\msvc64
@copy .\lib\x32\hbxlsxwriter.a \harbour\lib\win\msvc