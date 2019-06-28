##############################################################
# 祭日入りカレンダー CSV 作成
##############################################################
Param([string]$Year, $Path)

# 祭日データ
$TergetURI = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"

##############################################################
# オブジェクト作成
##############################################################
function AddWeek(){
	$Data = New-Object PSObject | Select-Object Col0,
												Col1,
												Col2,
												Col3,
												Col4,
												Col5,
												Col6
	return $Data
}

##############################################################
# 祭日データ取得
##############################################################
function GetSaijitsuData(){
	# Temp File
	$TempFile = Join-Path $env:TEMP syukujitsu.csv

	# 祝日 CSV ダウンロード
	Invoke-WebRequest -Uri $TergetURI -OutFile $TempFile

	# ヘッダー書き直し
	$text = Get-Content  -Path $TempFile -Encoding Oem
	$text[0] = "Date, Name"
	$text | Set-Content -Path $TempFile -Encoding Oem

	# CSV 読む
	$Lines = Import-Csv -Path $TempFile -Encoding Oem

	# Temp File 削除
	Remove-Item $TempFile

	return $Lines
}

##############################################################
# 祭日ハッシュテーブル作成
##############################################################
function GetSaijitsuHash([array]$Lines){
	# 祝日をハッシュテーブルに格納
	$SaijitsuHashTable = @{}
	foreach($Line in $Lines){
		$SaijitsuHashTable.Add( [datetime]$Line.Date, $Line.Name )
	}

	return $SaijitsuHashTable
}

##############################################################
# Main
##############################################################

if( $Year -eq [string]$null ){
	echo "Usage..."
	echo "    GetJapaneseCalendarCSV.ps1 -Year 対象年 -Path 出力先"
	exit
}

# 祭日データ取得
$Datas = GetSaijitsuData

# 祭日ハッシュテーブル作成
$SaijitsuHashTable = GetSaijitsuHash $Datas

# カレンダー
$Weeks = @()

# 開始日
try{
	$StartDay = Get-Date ($Year + "/1/1")
}
catch{
	echo "$Year は年として認識できません"
	exit
}

$i = 0
while($StartDay.AddDays($i).Year -eq [int]$Year){
	if( ($StartDay.AddDays($i).Day -eq 1) -and ($StartDay.AddDays($i).Month -eq 1) ){
		# 年表示
		$Week = AddWeek
		$Week.Col0 = ($StartDay.AddDays($i).Year).ToString() + " 年"
		$Weeks += $Week
		$Week = AddWeek
	}

	# 1日
	if($StartDay.AddDays($i).Day -eq 1 ){
		$Weeks += $Week

		$Week = AddWeek
		$Weeks += $Week

		$Week = AddWeek
		$Week.Col0 = ($StartDay.AddDays($i).Month).ToString() + " 月"
		$Weeks += $Week

		$Week = AddWeek
		$Week.Col0 = "日"
		$Week.Col1 = "月"
		$Week.Col2 = "火"
		$Week.Col3 = "水"
		$Week.Col4 = "木"
		$Week.Col5 = "金"
		$Week.Col6 = "土"
		$Weeks += $Week

		$Week = AddWeek
	}

	$Holiday = $SaijitsuHashTable[$StartDay.AddDays($i).Date]
	switch ($StartDay.AddDays($i).DayOfWeek){
		"Sunday" {
			$Week.Col0 = ($StartDay.AddDays($i).Day).ToString()
			if($Holiday -ne $null ){
				$Week.Col0 += "`r`n" + $Holiday
			}
		}

		"Monday" {
			$Week.Col1 = ($StartDay.AddDays($i).Day).ToString()
			if($Holiday -ne $null ){
				$Week.Col1 += "`r`n" + $Holiday
			}
		}

		"Tuesday" {
			$Week.Col2 = ($StartDay.AddDays($i).Day).ToString()
			if($Holiday -ne $null ){
				$Week.Col2 += "`r`n" + $Holiday
			}
		}

		"Wednesday" {
			$Week.Col3 = ($StartDay.AddDays($i).Day).ToString()
			if($Holiday -ne $null ){
				$Week.Col3 += "`r`n" + $Holiday
			}
		}

		"Thursday" {
			$Week.Col4 = ($StartDay.AddDays($i).Day).ToString()
			if($Holiday -ne $null ){
				$Week.Col4 += "`r`n" + $Holiday
			}
		}

		"Friday" {
			$Week.Col5 = ($StartDay.AddDays($i).Day).ToString()
			if($Holiday -ne $null ){
				$Week.Col5 += "`r`n" + $Holiday
			}
		}

		"Saturday" {
			$Week.Col6 = ($StartDay.AddDays($i).Day).ToString()
			if($Holiday -ne $null ){
				$Week.Col6 += "`r`n" + $Holiday
			}

			$Weeks += $Week
			$Week = AddWeek

		}
	}

	$i++
}

# CSV 出力
try{
	$Weeks | Export-CSV $Path -Encoding Oem
}
catch{
	echo "$Path を出力できませんでした"
	exit
}


