chcp 1251
@Set Pach1=C:\Users\SibNout2020\001_Лекци_VBA\
@Set Pach2=C:\D\Git_Hub\HowTo_VBA\
::copy "%Pach1%2ИСИП-319.xls" "%Pach2%2ИСИП-319.xls"
@copy "%Pach1%*.ipynb" "%Pach2%"
@copy "%Pach1%*.html" "%Pach2%"
::pause
TIMEOUT /T 5