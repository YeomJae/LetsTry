import xlsxDownload as xld
import xlsxProcess as xlp
import mailSend as ms

xld.xlsxDownload()

xlp.xlsxProcess('구매요구목록.xlsx','구매현황_일일보고.xlsx')

ms.SMTP_send('구매현황_일일보고.xlsx')