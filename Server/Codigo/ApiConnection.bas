Attribute VB_Name = "ApiConnection"
'**************************************************************
'Autor: Lucas Recoaro (Recox)
'Este modulo es el encargado de hacer los requests a los diferentes endpoints de la API
'La API esta escrita en javascript (Node.js/Express) y podemos desde obtener datos a hacer backup de charfiles o cuentas
'https://github.com/ao-libre/ao-api-server
'**************************************************************

Option Explicit
Private XmlHttp As Object
Private Endpoint As String
Private Parameters As String

Public Sub ApiEndpointBackupLogs()
    'Este endpoint hace una copia de todos los logs a una base de datos mysql
    
    Endpoint = ApiUrlServer & "/api/v1/logs/backuplogs"

    Call SendGETRequest(Endpoint)
End Sub

Public Sub ApiEndpointSendWelcomeEmail(ByVal UserName As String, ByVal Password As String, ByVal Email As String)
    'Este endpoint envia un email de bienvenida al usuario, con su nombre de usuario y password para que no lo pierda :)
    
    Endpoint = ApiUrlServer & "/api/v1/emails/welcome"
    
    Parameters = "username=" & UserName & "&password=" & Password & "&emailTo=" & Email

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendLoginAccountEmail(ByVal Email As String)
    'Este endpoint envia un email de login cada ves que un usuario se conecta
    Endpoint = ApiUrlServer & "/api/v1/emails/loginAccount"
    
    Parameters = "emailTo=" & Email

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Private Sub SendPOSTRequest(ByVal Endpoint As String, ByVal Parameters As String)

On Error GoTo ErrorHandler

    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "POST", Endpoint, True
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        
    'Por alguna razon tengo que castearlo a string, sino no funciona, la verdad no tengo idea por que ya que la variable es String
    XmlHttp.send CStr(Parameters)
        
        Set XmlHttp = Nothing

ErrorHandler:

    If Err.Number <> 0 Then
        Call LogError("Error POST endpoint: " & Endpoint & ". La Api parece estar offline. " & Err.Number & " - " & Err.description)
    End If
    
End Sub

Private Sub SendGETRequest(ByVal Endpoint As String)
On Error GoTo ErrorHandler

    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "GET", Endpoint, True
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send
        
        Set XmlHttp = Nothing

ErrorHandler:

    If Err.Number <> 0 Then
        Call LogError("Error GET endpoint: " & Endpoint & ". La Api parece estar offline. " & Err.Number & " - " & Err.description)
    End If
    
End Sub
