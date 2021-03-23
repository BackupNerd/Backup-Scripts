## Single Line Linux Install

CUID="6079722f-408c-473e-b991-aa57f4773b20"; PROFILE='115652'; INSTALL="swibm#$CUID#$PROFILE#.run" && curl -o $INSTALL https://cdn.cloudbackup.management/maxdownloads/mxb-linux-x86_64.run && chmod +x $INSTALL && sudo -s ./$INSTALL; rm ./swibm#*.run -f

## Single Line Linux Uninstall

sudo -s /opt/MXB/sbin/uninstall-fp.sh -s

## Single Line macOS Install

CUID="6079722f-408c-473e-b991-aa57f4773b20"; PROFILE='115652'; INSTALL="swibm#$CUID#$PROFILE#.pkg" && curl -o $INSTALL https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg && sudo installer -pkg $INSTALL -target /Applications; rm -f ./swibm#*.pkg

## Single Line Windows Install

$CUID="6079722f-408c-473e-b991-aa57f4773b20"; $PROFILE='115652'; $PRODUCT='All-In [60 Day]'; $INSTALL="c:\windows\temp\swibm#$CUID#$PROFILE#.exe"; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;(New-Object System.Net.WebClient).DownloadFile("http://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","$($INSTALL)"); & $INSTALL -product-name `"$PRODUCT`"
