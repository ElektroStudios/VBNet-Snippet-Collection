@Echo OFF & Title VB.NET Snippet Collection Installer by Elektro

Set "SnippetsDirectory=%UserProfile%\Documents\Visual Studio 2013\Code Snippets\Visual Basic\My Code Snippets"

If Not Exist "%SnippetsDirectory%\" (
	Echo Can't proceed.
	Echo Please set up correctly the 'SnippetsDirectory' location...
) Else (
	RD    /Q /S "%SnippetsDirectory%"
	XCopy /E /Y ".\*" "%SnippetsDirectory%\"
	DEL   /Q    "%SnippetsDirectory%\%~nx0"
)

Timeout /T 5