@echo off
"ExcelToJson.exe" in Localization.xlsx out mst_Data/StreamingAssets/Localization.json
"ExcelToJson.exe" in LocalizationAbility.xlsx out ms_Data/StreamingAssets/LocalizationAbility.json
"ExcelToJson.exe" in LocalizationArticle.xlsx out mst_Data/StreamingAssets/LocalizationArticle.json
"ExcelToJson.exe" in LocalizationBattle.xlsx out mst_Data/StreamingAssets/LocalizationBattle.json
"ExcelToJson.exe" in LocalizationJob.xlsx out mst_Data/StreamingAssets/LocalizationJob.json
"ExcelToJson.exe" in LocalizationCharacter.xlsx out mst_Data/StreamingAssets/LocalizationCharacter.json
"ExcelToJson.exe" in LocalizationUnit.xlsx out mst_Data/StreamingAssets/LocalizationUnit.json
"ExcelToJson.exe" in LocalizationElement.xlsx out mst_Data/StreamingAssets/LocalizationElement.json
"ExcelToJson.exe" in LocalizationScript.xlsx out mst_Data/StreamingAssets/LocalizationScript.json
"ExcelToJson.exe" in LocalizationCredit.xlsx out mst_Data/StreamingAssets/LocalizationCredit.json
echo finished all
pause >nul
