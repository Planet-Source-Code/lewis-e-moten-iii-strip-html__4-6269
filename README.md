<div align="center">

## Strip HTML


</div>

### Description

Strips any HTML tags from a string and returns the results. I use this for data that users submit through forms to prevent them from using HTML.
 
### More Info
 
asHTML - the string that may contain HTML code

This function assumes that you have vbScript 5.0 or higher installed.

Returns a string that has been stripped of HTML code.

Your page will come up with errors if you do not have vbScript 5.0 or higher installed on your server (or on the client browser if that is where it is being used)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Validation/ Processing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/validation-processing__4-16.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-strip-html__4-6269/archive/master.zip)

### API Declarations

Copyright (C) 1999, Lewis Moten. All rights reserved.


### Source Code

```
' Requires VBScript 5.0 installed on the server.
Function StripHTML(ByRef asHTML)
	Dim loRegExp	' Regular Expression Object
	' Create built in Regular Expression object
	Set loRegExp = New RegExp
	' Set the pattern to look for HTML tags
	loRegExp.Pattern = "<[^>]*>"
	' Return the original string stripped of HTML
	StripHTML = loRegExp.Replace(asHTML, "")
	' Release object from memory
	Set loRegExp = Nothing
End Function
```

