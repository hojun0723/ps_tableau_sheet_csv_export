#======================================================
#
# 월별 모니터링 결과 추출
# - tableau 통합문서의 데이터를 추출하여 csv파일로 저장
#
#======================================================

ECHO OFF
D:
cd "D:\Tableau Server\10.4\bin"
"***** Tableau Server Login"
tabcmd login -s http://localhost:8000 -u administrator -p password

$yyyyMMdd = Get-Date -format "yyyyMMdd"
$yyyy = $yyyyMMdd.Substring(0,4)
$MM = $yyyyMMdd.Substring(4,2)

#CSV화면의 기준월로 설정
$P_yyyyMMdd = Get-Date (Get-Date).AddMonths(-1) -format "yyyyMMdd"
$P_yyyy = $P_yyyyMMdd.Substring(0,4)
$P_MM = $P_yyyyMMdd.Substring(4,2)

$extension = "_0001.csv"

"***** ETL Processing"
$yyyyMMdd
tabcmd export "00_03__CSV_v5/FCSVdown?P_D_MONTH=$P_MM&P_D_YEAR=$P_yyyy&:embed=y&:showAppBanner=false&:showShareOptions=true&:display_count=no&:showVizHome=no" --csv -f "D:\dw\bi\BI_TOSBI001_$yyyyMMdd$extension"

"***** FILE Processing"
$Csv = Import-Csv -path "D:\dw\bi\BI_TOSBI001_$yyyyMMdd$extension" -Encoding UTF8

# 추출결과에서 ,를 제거

$Col = 'C_차원1'
$Csv | Foreach {$_.$Col = $_.$Col.Replace(",","")}

$Col = 'C_차원2'
$Csv | Foreach {$_.$Col = $_.$Col.Replace(",","")}

$Col = 'C_차원3'
$Csv | Foreach {$_.$Col = $_.$Col.Replace(",","")}

$Col = 'C_차원4'
$Csv | Foreach {$_.$Col = $_.$Col.Replace(",","")}

$Col = 'C_차원5'
$Csv | Foreach {$_.$Col = $_.$Col.Replace(",","")}

$Col = 'C_차원6'
$Csv | Foreach {$_.$Col = $_.$Col.Replace(",","")}

$Col = 'C_차원7'
$Csv | Foreach {$_.$Col = $_.$Col.Replace(",","")}

$CsvObject = $Csv | Convertto-Csv -NoTypeInformation
$CsvObject.replace('"','') | 
Set-Content "D:\dw\bi\BI_TOSBI001_$yyyyMMdd$extension" -Encoding UTF8