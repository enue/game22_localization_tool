@echo off
"ExcelToJson.exe" out Localization.xlsx in mst_Data/StreamingAssets/Localization.json
"ExcelToJson.exe" out LocalizationAbility.xlsx in mst_Data/StreamingAssets/LocalizationAbility.json
"ExcelToJson.exe" out LocalizationArticle.xlsx in mst_Data/StreamingAssets/LocalizationArticle.json
"ExcelToJson.exe" out LocalizationBattle.xlsx in mst_Data/StreamingAssets/LocalizationBattle.json
"ExcelToJson.exe" out LocalizationJob.xlsx in mst_Data/StreamingAssets/LocalizationJob.json
"ExcelToJson.exe" out LocalizationCharacter.xlsx in mst_Data/StreamingAssets/LocalizationCharacter.json
"ExcelToJson.exe" out LocalizationUnit.xlsx in mst_Data/StreamingAssets/LocalizationUnit.json
"ExcelToJson.exe" out LocalizationElement.xlsx in mst_Data/StreamingAssets/LocalizationElement.json
"ExcelToJson.exe" out LocalizationScript.xlsx in mst_Data/StreamingAssets/LocalizationScript.json
"ExcelToJson.exe" out LocalizationCredit.xlsx in mst_Data/StreamingAssets/LocalizationCredit.json
echo finished all
pause >nul
