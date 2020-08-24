# Powershell Archive Script Automation based on input CSV.
Multiple scripts to perform required tasks 
Powershell Archive Script with input as CSV

This Script can be used to archive files based on input days with retention days.

Just Schedule below command in Windows and you will get Daily Report of what archived and what left with failure reason if there are -
Archive_logs.ps1 -ArchiveConf Archive_logs_conf.csv 

CSV Input Can be defined as below -
1- UNCPath: Local/Network Share of directory where you want you apply archival/rotation action
2- ArchiveDays: How old files you want script to process and compress.
2- Retention: How many days of archived compressed files should be deleted permanently.
