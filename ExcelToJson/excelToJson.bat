@echo off
"ExcelToJson.exe" Localization.xlsx mst_develop_localize_Data/StreamingAssets/Localization.json
"ExcelToJson.exe" LocalizationAbility.xlsx mst_develop_localize_Data/StreamingAssets/LocalizationAbility.json
"ExcelToJson.exe" LocalizationArticle.xlsx mst_develop_localize_Data/StreamingAssets/LocalizationArticle.json
"ExcelToJson.exe" LocalizationBattle.xlsx mst_develop_localize_Data/StreamingAssets/LocalizationBattle.json
"ExcelToJson.exe" LocalizationJob.xlsx mst_develop_localize_Data/StreamingAssets/LocalizationJob.json
"ExcelToJson.exe" LocalizationCharacter.xlsx mst_develop_localize_Data/StreamingAssets/LocalizationCharacter.json
"ExcelToJson.exe" LocalizationUnit.xlsx mst_develop_localize_Data/StreamingAssets/LocalizationUnit.json
"ExcelToJson.exe" LocalizationElement.xlsx mst_develop_localize_Data/StreamingAssets/LocalizationElement.json
"ExcelToJson.exe" LocalizationScript.xlsx mst_develop_localize_Data/StreamingAssets/LocalizationScript.json
"ExcelToJson.exe" LocalizationCredit.xlsx mst_develop_localize_Data/StreamingAssets/LocalizationCredit.json
echo finished all
pause >nul