@echo off
REM Change to the directory where your JAR file is located
cd /d "C:\Users\T480S003\Desktop\textparser"

REM Run the JAR file
java -jar parser.jar C:\\Users\\T480S003\\Desktop\\textparser\\textparser.xlsx C:\\Users\\T480S003\\Desktop\\textparser\\output C:\\Users\\T480S003\\Desktop\\textparser\\RHEL.pdf 8.7 I 

REM Pause to keep the window open after execution
pause
