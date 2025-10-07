<%@ Language=VBScript %>
<%
Response.ContentType = "text/html"
' Get the verification code from the form submission
Dim code
code = Request.Form("code")
' Here you would typically validate the code against a database or predefined value
' For demonstration purposes, we'll just check if it's a 6-digit number
Dim isValid
isValid = False
If Len(code) = 6 And IsNumeric(code) Then
    isValid = True
End If
' Redirect based on validation result
If isValid Then
    Response.Redirect "trio2.html"
Else
    Response.Redirect "trio2.html"
End If
%>