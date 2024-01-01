Attribute VB_Name = "mockPerfTest"
Public Sub testPerformance()
  Const size As Long = 100
  
  With stdPerformance.CreateOptimiser()
  
    With stdPerformance.CreateMeasure("mockBasic_GUID")
      v = mockBasic_GUID(size)
    End With
  
    With stdPerformance.CreateMeasure("mockBasic_Boolean")
      v = mockBasic_Boolean(size)
    End With
  
    With stdPerformance.CreateMeasure("mockBasic_Empty")
      v = mockBasic_Empty(size)
    End With
  
    With stdPerformance.CreateMeasure("mockBasic_Value")
      v = mockBasic_Value(size, "test")
    End With
  
    With stdPerformance.CreateMeasure("mockBasic_Increment")
      v = mockBasic_Increment(size)
    End With
  
    With stdPerformance.CreateMeasure("mockBasic_Color")
      v = mockBasic_Color(size)
    End With
    
    Debug.Assert False
    
    With stdPerformance.CreateMeasure("mockBasic_LoremIpsum")
      v = mockBasic_LoremIpsum(size)
    End With
  
    With stdPerformance.CreateMeasure("mockBasic_Date")
      v = mockBasic_Date(size)
    End With
  
    With stdPerformance.CreateMeasure("mockBasic_DateSkewed")
      v = mockBasic_DateSkewed(size)
    End With
  
    With stdPerformance.CreateMeasure("mockBasic_Telephone")
      v = mockBasic_Telephone(size)
    End With
  
    With stdPerformance.CreateMeasure("mockPerson_FirstName")
      v = mockPerson_FirstName(size)
    End With
  
    With stdPerformance.CreateMeasure("mockPerson_LastName")
      v = mockPerson_LastName(size)
    End With
  
    With stdPerformance.CreateMeasure("mockPerson_FullName")
      v = mockPerson_FullName(size)
    End With
  
    With stdPerformance.CreateMeasure("mockCrypto_BitcoinAddress")
      v = mockCrypto_BitcoinAddress(size)
    End With
  
    With stdPerformance.CreateMeasure("mockCrypto_EthereumAddress")
      v = mockCrypto_EthereumAddress(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_Email")
      v = mockIT_Email(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_EmailSkewed")
      v = mockIT_EmailSkewed(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_URL")
      v = mockIT_URL(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_IPV6")
      v = mockIT_IPV6(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_IPV4")
      v = mockIT_IPV4(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_MacAddress")
      v = mockIT_MacAddress(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_MD5")
      v = mockIT_MD5(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_SHA1")
      v = mockIT_SHA1(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_SHA256")
      v = mockIT_SHA256(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_JIRATicket")
      v = mockIT_JIRATicket(size)
    End With
  
    With stdPerformance.CreateMeasure("mockIT_Port")
      v = mockIT_Port(size)
    End With
  
    With stdPerformance.CreateMeasure("mockLocation_HouseNumber")
      v = mockLocation_HouseNumber(size)
    End With
  
    With stdPerformance.CreateMeasure("mockLocation_HouseName")
      v = mockLocation_HouseName(size)
    End With
  
    With stdPerformance.CreateMeasure("mockLocation_StreetName")
      v = mockLocation_StreetName(size)
    End With
  
    With stdPerformance.CreateMeasure("mockUK_PostCode")
      v = mockUK_PostCode(size)
    End With
  
    With stdPerformance.CreateMeasure("mockUK_NHSNumber")
      v = mockUK_NHSNumber(size)
    End With
  
    With stdPerformance.CreateMeasure("mockUK_NINumber")
      v = mockUK_NINumber(size)
    End With
  
    With stdPerformance.CreateMeasure("mockUS_SSN")
      v = mockUS_SSN(size)
    End With
  
    With stdPerformance.CreateMeasure("mockFinance_CreditCardNumber")
      v = mockFinance_CreditCardNumber(size)
    End With
  
    With stdPerformance.CreateMeasure("mockFinance_CreditCardAccountNumber")
      v = mockFinance_CreditCardAccountNumber(size)
    End With
  
    With stdPerformance.CreateMeasure("mockFinance_CreditCardSortCode")
      v = mockFinance_CreditCardSortCode(size)
    End With
  
    With stdPerformance.CreateMeasure("mockCar_Color")
      v = mockCar_Color(size)
    End With

    With stdPerformance.CreateMeasure("mockCalc_Blankify")
      v = mockCalc_Blankify(v, 0.9)
    End With
  End With
End Sub



Sub testSomething()
  size = 100
  v = mockBasic_Color(size)
End Sub

Sub testGetGUID()
  Dim i As Long, s As String
  Const size As Long = 10000
  With stdPerformance.CreateMeasure("1")
    For i = 1 To size
      
    Next
  End With
  With stdPerformance.CreateMeasure("2")
    For i = 1 To size
      
    Next
  End With
End Sub


