SET CONVERTER=%ExcelToEnum_Root%\Executables\ExcelToEnum.exe
SET SCENARIO_DIR=%ExcelToEnum_Root%\Data\\
SET OUTPUT_DIR=%ExcelToEnum_Root%\\
SET SETTINGS_DIR=%ExcelToEnum_Root%\Settings\\
SET NAMESPACE=JK

"%CONVERTER%" "%SCENARIO_DIR%" "%OUTPUT_DIR%" "%SETTINGS_DIR%" "%NAMESPACE%"
