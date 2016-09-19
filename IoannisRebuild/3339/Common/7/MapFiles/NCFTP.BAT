
CD C:\
copy /b lb*.txt +,,


CD C:\Touchsource\ASP\#
"C:\Program Files\NcFTP\ncftpget.exe" -u # -p # -DD ftp://ftp.touchsource.com/lb*-@.txt
"C:\Program Files\NcFTP\ncftpget.exe" -u # -p # ftp://ftp.touchsource.com/*.*
"C:\Program Files\NcFTP\ncftpget.exe" -u # -p # ftp://ftp.touchsource.com/Common/@/MapFiles/*.*


CD C:\Touchsource\ASP\#\Common\Program
"C:\Program Files\NcFTP\ncftpget.exe" -u # -p # ftp://ftp.touchsource.com/common/program/*.*

CD C:\Touchsource\ASP\#\Common\@
"C:\Program Files\NcFTP\ncftpget.exe" -u # -p # ftp://ftp.touchsource.com/Common/@/*.*

CD "C:\Users\Administrator\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
"C:\Program Files\NcFTP\ncftpget.exe" -u # -p # ftp://ftp.touchsource.com/Startup/*.*

"C:\Program Files\NcFTP\ncftpput.exe" -u # -p # ftp.touchsource.com / c:\lb*.txt

DEL "C:\Users\Administrator\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Thumbs.db"
DEL "C:\Users\Administrator\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\lb*.txt"
DEL C:\Touchsource\ASP\#\lb*.txt
EXIT




