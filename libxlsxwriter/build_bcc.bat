@del .\lib\x64\xlsxwriter.a
@del .\lib\x32\xlsxwriter.a
@hbmk2 libxlsxwriter.hbp -rebuildall
@copy .\lib\x64\xlsxwriter.a \harbour\lib\win\bcc64
@copy .\lib\x32\xlsxwriter.a \harbour\lib\win\bcc