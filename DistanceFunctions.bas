Attribute VB_Name = "DistanceFunctions"
Option Explicit

'Declare the Sleep function from the Windows DLL
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Public Function GetDistance(origins As String, destinations As String) As String
    'Written By:    Santiago Delgado
    'Last Updated:  10/12/2020
    'Based on:      https://myengineeringworld.net/2014/06/geocoding-vba-google-api.html
    
    
    'Declaring the necessary variables.
    Dim apiKey              As String
    Dim xmlhttpRequest      As Object
    Dim xmlDoc              As Object
    Dim xmlStatusNode       As Object
    Dim xmlLatitudeNode     As Object
    Dim xmLongitudeNode     As Object
    Dim xmlResultNode       As Object
    Dim xmlPath             As String
       
    'Set your API key in this variable. Check this link for more info:
    apiKey = "INSERT-YOUR-API-KEY-HERE"
    
    'Check that an API key has been provided.
    If apiKey = vbNullString Or apiKey = "The API Key" Then
        GetDistance = "Empty or invalid API Key"
        Exit Function
    End If
    
    'Generic error handling.
    On Error GoTo errorHandler
            
    'Create the request object and check if it was created successfully.
    Set xmlhttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
    If xmlhttpRequest Is Nothing Then
        GetDistance = "Cannot create the request object"
        Exit Function
    End If
        
    'Create the request based on Google Geocoding API. Parameters (from Google page):
    '- Address: The address that you want to geocode.
    
    'Note: The EncodeURL function was added to allow users from Greece, Poland, Germany, France and other countries
    'geocode address from their home countries without a problem. The particular function (EncodeURL),
    'returns a URL-encoded string without the special characters.
    'This function, however, was introduced in Excel 2013, so it will NOT work in older Excel versions.
    xmlhttpRequest.Open "GET", "https://maps.googleapis.com/maps/api/distancematrix/xml?units=imperial" _
    & "&origins=" & Application.EncodeURL(origins) & "&destinations=" & Application.EncodeURL(destinations) & "&key=" & apiKey, False
    
    'An alternative way, without the EncodeURL function, will be this:
    'xmlhttpRequest.Open "GET", "https://maps.googleapis.com/maps/api/geocode/xml?" & "&address=" & Address & "&key=" & ApiKey, False
    
    'Wait ten seconds to not overload the API
    Sleep (10000)
    
    'Send the request to the Google server.
    xmlhttpRequest.send
    
    'Create the DOM document object and check if it was created successfully.
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    If xmlDoc Is Nothing Then
        GetDistance = "Cannot create the DOM document object"
        Exit Function
    End If
    
    'Read the XML results from the request.
    xmlDoc.LoadXML xmlhttpRequest.responseText
    
    'Get the value from the status node.
    Set xmlStatusNode = xmlDoc.SelectSingleNode("//DistanceMatrixResponse/row/element/status")
    
    'Based on the status node result, proceed accordingly.
    Select Case UCase(xmlStatusNode.Text)
    
        Case "OK"                       'The API request was successful.
                                        'At least one result was returned.
            xmlPath = "//DistanceMatrixResponse/row/element/distance/value"
            Set xmlResultNode = xmlDoc.SelectSingleNode(xmlPath)
            GetDistance = xmlResultNode.Text
        
        Case "ZERO_RESULTS"             'The geocode was successful but returned no results.
            GetDistance = "No results were found"
            
        Case "OVER_DAILY_LIMIT"         'Indicates any of the following:
                                        '- The API key is missing or invalid.
                                        '- Billing has not been enabled on your account.
                                        '- A self-imposed usage cap has been exceeded.
                                        '- The provided method of payment is no longer valid
                                        '  (for example, a credit card has expired).
            GetDistance = "Rate, billing, or payment problem"
            
        Case "OVER_QUERY_LIMIT"         'The requestor has exceeded the quota limit.
            GetDistance = "Quota limit exceeded"
            
        Case "REQUEST_DENIED"           'The API did not complete the request.
            GetDistance = "Server denied the request"
            
        Case "INVALID_REQUEST"           'The API request is empty or is malformed.
            GetDistance = "Request was empty or malformed"
        
        Case "UNKNOWN_ERROR"            'The request could not be processed due to a server error.
            GetDistance = "Unknown error"
        
        Case Else   'Just in case...
            GetDistance = "Error"
        
    End Select
        
    'Release the objects before exiting (or in case of error).
errorHandler:
    Set xmlStatusNode = Nothing
    Set xmlLatitudeNode = Nothing
    Set xmLongitudeNode = Nothing
    Set xmlDoc = Nothing
    Set xmlhttpRequest = Nothing
    
End Function



Public Function GetPostalcode(address As String) As String
    'Written By:    Santiago Delgado
    'Last Updated:  10/11/2020
    'Based on:      https://myengineeringworld.net/2014/06/geocoding-vba-google-api.html
    
    
    'Declaring the necessary variables.
    Dim apiKey              As String
    Dim xmlhttpRequest      As Object
    Dim xmlDoc              As Object
    Dim xmlStatusNode       As Object
    Dim xmlLatitudeNode     As Object
    Dim xmLongitudeNode     As Object
    Dim xmlPostalCodeNode   As Object
    Dim xmlPath             As String
       
    'Set your API key in this variable. Check this link for more info:
    apiKey = "INSERT-YOUR-API-KEY-HERE"
    
    'Check that an API key has been provided.
    If apiKey = vbNullString Or apiKey = "The API Key" Then
        GetPostalcode = "Empty or invalid API Key"
        Exit Function
    End If
    
    'Generic error handling.
    On Error GoTo errorHandler
            
    'Create the request object and check if it was created successfully.
    Set xmlhttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
    If xmlhttpRequest Is Nothing Then
        GetPostalcode = "Cannot create the request object"
        Exit Function
    End If
        
    'Create the request based on Google Geocoding API. Parameters (from Google page):
    '- Address: The address that you want to geocode.
    
    'Note: The EncodeURL function was added to allow users from Greece, Poland, Germany, France and other countries
    'geocode address from their home countries without a problem. The particular function (EncodeURL),
    'returns a URL-encoded string without the special characters.
    'This function, however, was introduced in Excel 2013, so it will NOT work in older Excel versions.
    xmlhttpRequest.Open "GET", "https://maps.googleapis.com/maps/api/geocode/xml?" _
    & "&address=" & Application.EncodeURL(address) & "&key=" & apiKey, False
    
    'An alternative way, without the EncodeURL function, will be this:
    'xmlhttpRequest.Open "GET", "https://maps.googleapis.com/maps/api/geocode/xml?" & "&address=" & Address & "&key=" & ApiKey, False
    
    'Send the request to the Google server.
    xmlhttpRequest.send
    
    'Wait ten seconds to not overload the API
    Sleep (10000)
    
    'Create the DOM document object and check if it was created successfully.
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    If xmlDoc Is Nothing Then
        GetPostalcode = "Cannot create the DOM document object"
        Exit Function
    End If
    
    'Read the XML results from the request.
    xmlDoc.LoadXML xmlhttpRequest.responseText
    
    'Get the value from the status node.
    Set xmlStatusNode = xmlDoc.SelectSingleNode("//status")
    
    'Based on the status node result, proceed accordingly.
    Select Case UCase(xmlStatusNode.Text)
    
        Case "OK"                       'The API request was successful.
                                        'At least one result was returned.
            xmlPath = "//result/address_component[type='postal_code']/short_name"
            Set xmlPostalCodeNode = xmlDoc.SelectSingleNode(xmlPath)
            GetPostalcode = xmlPostalCodeNode.Text
        
        Case "ZERO_RESULTS"             'The geocode was successful but returned no results.
            GetPostalcode = "No results were found"
            
        Case "OVER_DAILY_LIMIT"         'Indicates any of the following:
                                        '- The API key is missing or invalid.
                                        '- Billing has not been enabled on your account.
                                        '- A self-imposed usage cap has been exceeded.
                                        '- The provided method of payment is no longer valid
                                        '  (for example, a credit card has expired).
            GetPostalcode = "Rate, billing, or payment problem"
            
        Case "OVER_QUERY_LIMIT"         'The requestor has exceeded the quota limit.
            GetPostalcode = "Quota limit exceeded"
            
        Case "REQUEST_DENIED"           'The API did not complete the request.
            GetPostalcode = "Server denied the request"
            
        Case "INVALID_REQUEST"           'The API request is empty or is malformed.
            GetPostalcode = "Request was empty or malformed"
        
        Case "UNKNOWN_ERROR"            'The request could not be processed due to a server error.
            GetPostalcode = "Unknown error"
        
        Case Else   'Just in case...
            GetPostalcode = "Error"
        
    End Select
        
    'Release the objects before exiting (or in case of error).
errorHandler:
    Set xmlStatusNode = Nothing
    Set xmlLatitudeNode = Nothing
    Set xmLongitudeNode = Nothing
    Set xmlDoc = Nothing
    Set xmlhttpRequest = Nothing
    
End Function

Public Function GetAddress(address As String) As String
    'Written By:    Santiago Delgado
    'Last Updated:  10/11/2020
    'Based on:      https://myengineeringworld.net/2014/06/geocoding-vba-google-api.html
    
    
    'Declaring the necessary variables.
    Dim apiKey              As String
    Dim xmlhttpRequest      As Object
    Dim xmlDoc              As Object
    Dim xmlStatusNode       As Object
    Dim xmlLatitudeNode     As Object
    Dim xmLongitudeNode     As Object
    Dim xmlAddressNode      As Object
    Dim xmlPath             As String
       
    'Set your API key in this variable. Check this link for more info:
    apiKey = "INSERT-YOUR-API-KEY-HERE"
    
    'Check that an API key has been provided.
    If apiKey = vbNullString Or apiKey = "The API Key" Then
        GetAddress = "Empty or invalid API Key"
        Exit Function
    End If
    
    'Generic error handling.
    On Error GoTo errorHandler
            
    'Create the request object and check if it was created successfully.
    Set xmlhttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
    If xmlhttpRequest Is Nothing Then
        GetAddress = "Cannot create the request object"
        Exit Function
    End If
        
    'Create the request based on Google Geocoding API. Parameters (from Google page):
    '- Address: The address that you want to geocode.
    
    'Note: The EncodeURL function was added to allow users from Greece, Poland, Germany, France and other countries
    'geocode address from their home countries without a problem. The particular function (EncodeURL),
    'returns a URL-encoded string without the special characters.
    'This function, however, was introduced in Excel 2013, so it will NOT work in older Excel versions.
    xmlhttpRequest.Open "GET", "https://maps.googleapis.com/maps/api/geocode/xml?" _
    & "&address=" & Application.EncodeURL(address) & "&key=" & apiKey, False
    
    'An alternative way, without the EncodeURL function, will be this:
    'xmlhttpRequest.Open "GET", "https://maps.googleapis.com/maps/api/geocode/xml?" & "&address=" & Address & "&key=" & ApiKey, False
    
    'Send the request to the Google server.
    xmlhttpRequest.send
    
    'Wait ten seconds to not overload the API
    Sleep (10000)
    
    'Create the DOM document object and check if it was created successfully.
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    If xmlDoc Is Nothing Then
        GetAddress = "Cannot create the DOM document object"
        Exit Function
    End If
    
    'Read the XML results from the request.
    xmlDoc.LoadXML xmlhttpRequest.responseText
    
    'Get the value from the status node.
    Set xmlStatusNode = xmlDoc.SelectSingleNode("//status")
    
    'Based on the status node result, proceed accordingly.
    Select Case UCase(xmlStatusNode.Text)
    
        Case "OK"                       'The API request was successful.
                                        'At least one result was returned.
            xmlPath = "//result/formatted_address"
            Set xmlAddressNode = xmlDoc.SelectSingleNode(xmlPath)
            GetAddress = xmlAddressNode.Text
        
        Case "ZERO_RESULTS"             'The geocode was successful but returned no results.
            GetAddress = "No results were found"
            
        Case "OVER_DAILY_LIMIT"         'Indicates any of the following:
                                        '- The API key is missing or invalid.
                                        '- Billing has not been enabled on your account.
                                        '- A self-imposed usage cap has been exceeded.
                                        '- The provided method of payment is no longer valid
                                        '  (for example, a credit card has expired).
            GetAddress = "Rate, billing, or payment problem"
            
        Case "OVER_QUERY_LIMIT"         'The requestor has exceeded the quota limit.
            GetAddress = "Quota limit exceeded"
            
        Case "REQUEST_DENIED"           'The API did not complete the request.
            GetAddress = "Server denied the request"
            
        Case "INVALID_REQUEST"           'The API request is empty or is malformed.
            GetAddress = "Request was empty or malformed"
        
        Case "UNKNOWN_ERROR"            'The request could not be processed due to a server error.
            GetAddress = "Unknown error"
        
        Case Else   'Just in case...
            GetAddress = "Error"
        
    End Select
        
    'Release the objects before exiting (or in case of error).
errorHandler:
    Set xmlStatusNode = Nothing
    Set xmlLatitudeNode = Nothing
    Set xmLongitudeNode = Nothing
    Set xmlDoc = Nothing
    Set xmlhttpRequest = Nothing
    
End Function





