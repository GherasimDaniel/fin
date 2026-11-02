@echo off
rem Starts the local helper and opens the month delta in your browser.
rem Each time the page opens it checks Yahoo for new trading days and keeps one month of history.
cd /d "%~dp0"
node serve.js
