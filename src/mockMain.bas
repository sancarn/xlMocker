Attribute VB_Name = "mockMain"

'Generate a random GUID
'@param iNumber - The number of items to generate
'@returns - A column of random GUIDs
Public Function mockBasic_GUID(ByVal iNumber As Long) As Variant
  mockBasic_GUID = genColumn(iNumber, stdCallback.CreateFromPointer(AddressOf getGUID, vbString))
End Function

'Generates a random boolean
'@param iNumber - The number of items to generate
'@param trueWeight - The probability of generating a true value
'@returns - A column of random booleans
Public Function mockBasic_Boolean(ByVal iNumber as Long, Optional ByVal trueWeight as double) As Variant
  mockBasic_Boolean = mockCalc_WeightedArray(iNumber, Array(True, trueWeight, False, 1 - trueWeight))
End Function

'Generate a column of empty values
'@param iNumber - The number of items to generate
'@returns - A column of `Empty` values
Public Function mockBasic_Empty(ByVal iNumber as Long) As Variant
  mockBasic_Empty = genStatic(iNumber, Empty)
End Function

'Generate a column containing a value
'@param iNumber - The number of items to generate
'@param value - The value in each cell of the column
'@returns - A column of the given value
Public Function mockBasic_Value(ByVal iNumber as Long, ByVal value as Variant) As Variant
  mockBasic_Error = genStatic(iNumber, value)
End Function

'Generate a column containing a random color
'@param iNumber - The number of items to generate
'@returns - A column of random colors
Public Function mockBasic_Color(ByVal iNumber as Long) as Variant
  mockBasic_Color = mockCalc_ValueFromRange(iNumber, BasicData.Range("Basic_Color[Color]"))
End Function

'Generate a column containing a random string of lorem ipsum text
'@param iNumber - The number of items to generate
'@param iMaxParagraphs - The maximum number of paragraphs to generate
'@param iMinParagraphs - The minimum number of paragraphs to generate
'@param iMaxSentences - The maximum number of sentences to generate
'@param iMinSentences - The minimum number of sentences to generate
'@param iMaxWords - The maximum number of words to generate
'@param iMinWords - The minimum number of words to generate
'@returns - A column of random lorem ipsum text
Public Function mockBasic_LoremIpsum(ByVal iNumber as Long, Optional ByVal iMaxParagraphs As Long = 1, Optional ByVal iMinParagraphs As Long = 1, Optional ByVal iMaxSentences As Long = 4, Optional ByVal iMinSentences As Long = 2, Optional ByVal iMaxWords As Long = 10, Optional ByVal iMinWords As Long = 5) As Variant
  mockBasic_LoremIpsum = genColumn(iNumber, stdCallback.CreateFromPointer(AddressOf generateLoremIpsum, vbString, Array(vbLong, vbLong, vbLong, vbLong, vbLong, vbLong)).Bind(iMaxParagraphs, iMinParagraphs, iMaxSentences, iMinSentences, iMaxWords, iMinWords))
End Function

'Generate a column of random dates
'@param iNumber - The number of items to generate
'@param iMaxDate - The maximum date
'@param iMinDate - The minimum date
'@returns - A column of random dates
Public Function mockBasic_Date(ByVal iNumber as Long, Optional ByVal iMaxDate As Date = 0, Optional ByVal iMinDate As Date = #1/1/2000#) As Variant
  if iMaxDate = 0 then iMaxDate = now()
  mockBasic_Date = genColumn(iNumber, stdCallback.CreateFromPointer(AddressOf genDate, vbDate, Array(vbDate, vbDate)).Bind(iMaxDate, iMinDate))
End Function

'Generate a column of random dates skewed with data quality issues as if it were written by a human
'@param iNumber - The number of items to generate
'@param iMaxDate - The maximum date
'@param iMinDate - The minimum date
'@returns - A column of random dates skewed with data quality issues as if it were written by a human
Public Function mockBasic_DateSkewed(ByVal iNumber as Long, Optional ByVal iMaxDate As Date = 0, Optional ByVal iMinDate As Date = #1/1/2000#) As Variant
  if iMaxDate = 0 then iMaxDate = now()
  mockBasic_DateSkewed = genColumn(iNumber, stdCallback.CreateFromPointer(AddressOf generateDateSkewed, vbDate, Array(vbDate, vbDate)).Bind(iMaxDate, iMinDate))
End Function

'Generate a column of random phone numbers
'@param iNumber - The number of items to generate
'@returns - A column of random phone numbers
'@remark - Original source: https://regex101.com/r/wZ4uU6/2
Public Function mockBasic_Telephone(ByVal iNumber as Long) as Variant
  mockBasic_Telephone = mockCalc_Regex(iNumber, "(?:([+]\d{1,4})[-.\s]?)?(?:[(](\d{1,3})[)][-.\s]?)?(\d{1,4})[-.\s]?(\d{1,4})[-.\s]?(\d{1,9})")
End Function

'Generate a column of random bitcoin addresses
'@param iNumber - The number of items to generate
'@returns - A column of random bitcoin addresses
'@remark - Original source: https://ihateregex.io/expr/bitcoin-address/
Public Function mockCrypto_BitcoinAddress(ByVal iNumber as Long) As Variant
  mockCrypto_BitcoinAddress = mockCalc_Regex(iNumber, "^(bc1|[13])[a-zA-HJ-NP-Z0-9]{25,39}$")
End Function

'Generate a column of random ethereum addresses
'@param iNumber - The number of items to generate
'@returns - A column of random ethereum addresses
'@remark - Original source: https://stackoverflow.com/questions/49451874/regex-to-match-string-containing-two-eth-address-in-any-order
Public Function mockCrypto_EthereumAddress(ByVal iNumber as Long) As Variant
  mockCrypto_EthereumAddress = mockCalc_Regex(iNumber, "(0x)?[0-9a-fA-F]{40}")
End Function

'Generate a column of random email addresses
'@param iNumber - The number of items to generate
'@returns - A column of random email addresses
'@remark - Original source: https://regex101.com/library/sI6yF5. Emails which comply with RFC2822.
Public Function mockIT_Email(ByVal iNumber as Long) As Variant
  mockIT_Email = mockCalc_Regex(iNumber, "([^\x00-\x20\x22\x28\x29\x2c\x2e\x3a-\x3c\x3e\x40\x5b-\x5d\x7f-\xff]+|\x22([^\x0d\x22\x5c\x80-\xff]|\x5c[\x00-\x7f])*\x22)(\x2e([^\x00-\x20\x22\x28\x29\x2c\x2e\x3a-\x3c\x3e\x40\x5b-\x5d\x7f-\xff]+|\x22([^\x0d\x22\x5c\x80-\xff]|\x5c[\x00-\x7f])*\x22))*\x40([^\x00-\x20\x22\x28\x29\x2c\x2e\x3a-\x3c\x3e\x40\x5b-\x5d\x7f-\xff]+|\x5b([^\x0d\x5b-\x5d\x80-\xff]|\x5c[\x00-\x7f])*\x5d)(\x2e([^\x00-\x20\x22\x28\x29\x2c\x2e\x3a-\x3c\x3e\x40\x5b-\x5d\x7f-\xff]+|\x5b([^\x0d\x5b-\x5d\x80-\xff]|\x5c[\x00-\x7f])*\x5d))*")
End Function

'Generate a column of random email addresses which may be distorted / contain errors
'@param iNumber - The number of items to generate
'@returns - A column of random email addresses
'@remark - Original source: https://regex101.com/r/wB7xJ7/1
Public Function mockIT_EmailSkewed(ByVal iNumber as Long) As Variant
  mockIT_EmailSkewed = mockCalc_Regex(iNumber, "^(?<Username>[-\w\d\.]+?)(?:\s+at\s+|\s*@\s*|\s*(?:[\[\]@]){3}\s*)(?<Domain>[-\w\d\.]*?)\s*(?:dot|\.|(?:[\[\]dot\.]){3,5})\s*(?<TLD>\w+)$")
End Function

'Generate a column of random URLs
'@param iNumber - The number of items to generate
'@returns - A column of random URLs
'@reamrk - Original source: https://ihateregex.io/expr/url/
Public Function mockIT_URL(ByVal iNumber as Long) As Variant
  mockIT_URL = mockCalc_Regex(iNumber, "https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()!@:%_\+.~#?&\/\/=]*)")
End Function

'Generate a column of random IPV6 addresses
'@param iNumber - The number of items to generate
'@returns - A column of random IPV6 addresses
'@remark - Original source: https://ihateregex.io/expr/ipv6/
Public Function mockIT_IPV6(ByVal iNumber as Long) As Variant
  mockIT_IPV6 = mockCalc_Regex(iNumber, "(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))")
End Function

'Generate a column of random IPV4 addresses
'@param iNumber - The number of items to generate
'@returns - A column of random IPV4 addresses
'@remark - Original source: https://ihateregex.io/expr/ip/
Public Function mockIT_IPV4(ByVal iNumber as Long) As Variant
  mockIT_IPV4 = mockCalc_Regex(iNumber, "(\b25[0-5]|\b2[0-4][0-9]|\b[01]?[0-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){3}")
End Function

'Generate a column of random mac addresses
'@param iNumber - The number of items to generate
'@returns - A column of random mac addresses
'@remark - Original source: https://ihateregex.io/expr/mac-address/
Public Function mockIT_MacAddress(ByVal iNumber as Long) As Variant
  mockIT_MacAddress = mockCalc_Regex(iNumber, "^[a-fA-F0-9]{2}(:[a-fA-F0-9]{2}){5}$")
End Function

'Generate a column of random MD5 hashes
'@param iNumber - The number of items to generate
'@returns - A column of random MD5 hashes
Public Function mockIT_MD5(ByVal iNumber as Long) As Variant
  mockIT_MD5 = mockCalc_Regex(iNumber, "[a-f0-9]{32}")
End Function

'Generate a column of random SHA1 hashes
'@param iNumber - The number of items to generate
'@returns - A column of random SHA1 hashes
Public Function mockIT_SHA1(ByVal iNumber as Long) As Variant
  mockIT_SHA1 = mockCalc_Regex(iNumber, "[a-f0-9]{40}")
End Function

'Generate a column of random SHA256 hashes
'@param iNumber - The number of items to generate
'@returns - A column of random SHA256 hashes
Public Function mockIT_SHA256(ByVal iNumber as Long) As Variant
  mockIT_SHA256 = mockCalc_Regex(iNumber, "[a-f0-9]{64}")
End Function

'Generate a column of random JIRA Ticket IDs
'@param iNumber - The number of items to generate
'@returns - A column of random JIRA Ticket IDs
'@remark - Original source: https://ihateregex.io/expr/jira-ticket/
Public Function mockIT_JIRATicket(ByVal iNumber as Long) As Variant
  mockIT_JIRATicket = mockCalc_Regex(iNumber, "[A-Z]{2,}-\d+")
End Function

'Generate a column of random Port Numbers
'@param iNumber - The number of items to generate
'@returns - A column of random port numbers
'@remark - Original source: https://ihateregex.io/expr/port/
Public Function mockIT_Port(ByVal iNumber as Long) As Variant
  mockIT_Port = mockCalc_Regex(iNumber, "^((6553[0-5])|(655[0-2][0-9])|(65[0-4][0-9]{2})|(6[0-4][0-9]{3})|([1-5][0-9]{4})|([0-5]{0,5})|([0-9]{1,4}))$")
End Function

'Generate a random house number
'@param iNumber - The number of items to generate
'@returns - A column of random house numbers
Public Function mockLocation_HouseNumber(ByVal iNumber as Long) As Variant
  mockLocation_HouseNumber = mockCalc_Regex(iNumber, "[1-9]\d{0,3}[a-d]?")
End Function

'Generate a random house name
'@param iNumber - The number of items to generate
'@returns - A column of random house names
Public Function mockLocation_HouseName(ByVal iNumber as Long) As Variant
  static houseNouns as variant: if isEmpty(houseNouns) then houseNouns = mockCalc_ValueFromRange(iNumber, BasicData.Range("HouseNouns[Word]"))
  static adjectives as variant: if isEmpty(adjectives) then adjectives = mockCalc_ValueFromRange(iNumber, BasicData.Range("Adjectives[Adjective]"))
  static nouns as variant: if isEmpty(nouns) then nouns = mockCalc_ValueFromRange(iNumber, BasicData.Range("Nouns[Noun]"))
  Dim vRet(): Redim vRet(1 To iNumber, 1 To 1)
  For i = 1 To iNumber
    select case rnd()
      case is < 0.1
        vRet(i, 1) = adjectives(i, 1) & " " & nouns(i, 1) & " " & houseNouns(i, 1)
      case is < 0.6
        vRet(i, 1) = nouns(i, 1) & " " & houseNouns(i, 1)
      case else
        vRet(i, 1) = adjectives(i, 1) & " " & houseNouns(i, 1)
    end select
    vRet(i, 1) = ProperCase(vRet(i, 1))
  next
  mockLocation_HouseName = vRet
End Function

'Generate a random street name
'@param iNumber - The number of items to generate
'@returns - A column of random street names
Public Function mockLocation_StreetName(ByVal iNumber as Long) As Variant
  static streetNouns as variant: if isEmpty(streetNouns) then streetNouns = mockCalc_ValueFromRange(iNumber, BasicData.Range("StreetNouns[Word]"))
  static adjectives as variant: if isEmpty(adjectives) then adjectives = mockCalc_ValueFromRange(iNumber, BasicData.Range("Adjectives[Adjective]"))
  static nouns as variant: if isEmpty(nouns) then nouns = mockCalc_ValueFromRange(iNumber, BasicData.Range("Nouns[Noun]"))
  Dim vRet(): Redim vRet(1 To iNumber, 1 To 1)
  For i = 1 To iNumber
    select case rnd()
      case is < 0.1
        vRet(i, 1) = adjectives(i, 1) & " " & nouns(i, 1) & " " & streetNouns(i, 1)
      case is < 0.6
        vRet(i, 1) = nouns(i, 1) & " " & streetNouns(i, 1)
      case else
        vRet(i, 1) = adjectives(i, 1) & " " & streetNouns(i, 1)
    end select
    vRet(i, 1) = ProperCase(vRet(i, 1))
  next
  mockLocation_StreetName = vRet
End Function

'Generate a random UK postcode
'@param iNumber - The number of items to generate
'@returns - A column of random postcodes
Public Function mockUK_PostCode(ByVal iNumber As Long) As Variant
  mockUK_PostCode = mockCalc_Regex(iNumber, "[A-Z]{1,2}\d{1,2} \d[A-Z]{2}")
End Function

'Make a random percentage of values blank
'@param vArr - The array to blankify
'@param percentBlank - The percentage of values to be blank
'@returns - The blankified array
Public Function Blankify(Byval vArr as variant, ByVal percentBlank as Double) as variant
  if percentBlank > 1 then percentBlank = 1
  if percentBlank < 0 then percentBlank = 0
  if percentBlank = 0 then Blankify = vArr: exit function

  Call Randomize
  dim i as long: For i = lbound(vArr,1) to ubound(vArr,1)
    if rnd() < percentBlank then
      vArr(i,1) = Empty
    end if
  next
  Blankify = vArr
End Function

'Generate a random UK NHS Number
'@param iNumber - The number of items to generate
'@returns - A column of random NHS numbers
Public Function mockUK_NHSNumber(ByVal iNumber As Long) As Variant
  mockUK_NHSNumber = mockCalc_Regex(iNumber, "\b\d{3}-\d{3}-\d{4}\b ")
End Function

'Generate a random UK National Insurance Number
'@param iNumber - The number of items to generate
'@returns - A column of random National Insurance numbers
'@remark Note: Due to current limitations with stdRegex, some invalid NI numbers may be generated due to lack of lookaround support in generation
Public Function mockUK_NINumber(ByVal iNumber As Long) As Variant
  mockUK_NINumber = mockCalc_Regex(iNumber, "(?!BG|GB|NK|KN|TN|NT|ZZ)[A-CEGHJ-PR-TW-Z][A-CEGHJ-NPR-TW-Z](?:\s?\d){6}\s?[A-D]")
End Function

'Generate a random USA Social Security Number
'@param iNumber - The number of items to generate
'@returns - A column of random SSN numbers
'@remark - original source: https://ihateregex.io/expr/ssn/
Public Function mockUS_SSN(ByVal iNumber As Long) As Variant
  mockUS_SSN = mockCalc_Regex(iNumber, "^(?!0{3})(?!6{3})[0-8]\d{2}-(?!0{2})\d{2}-(?!0{4})\d{4}$")
End Function

'Generate a random Credit Card Number
'@param iNumber - The number of items to generate
'@returns - A column of random credit card numbers
Public Function mockFinance_CreditCardNumber(ByVal iNumber As Long) As Variant
  mockFinance_CreditCardNumber = mockCalc_Regex(iNumber, "5[1-9]\d{2} \d{4} \d{4} \d{4}")
End Function

'Generate a random Credit Card Account Number
'@param iNumber - The number of items to generate
'@returns - A column of random account numbers
Public Function mockFinance_CreditCardAccountNumber(ByVal iNumber As Long) As Variant
  mockFinance_CreditCardAccountNumber = mockCalc_Regex(iNumber, "\d{8}")
End Function

'Generate a random Credit Card Sort Code
'@param iNumber - The number of items to generate
'@returns - A column of random sort codes with realistic distribution
Public Function mockFinance_CreditCardSortCode(ByVal iNumber As Long) As Variant
  mockFinance_CreditCardSortCode = mockCalc_ValueFromRangeWeighted(iNumber, FinData.Range("Fin_SortCode[Sort Code]"), FinData.Range("Fin_SortCode[Weight]"))
End Function

'Obtain a random car color weighted by popularity
'@param iNumber - The number of items to generate
'@returns - A column of random car colors with realistic distribution
'@remark Car color weights found here: https://www.carpro.com/blog/most-popular-2023-model-year-car-colors
Public Function mockCar_Color(ByVal iNumber As Long) As Variant
  mockCar_Color = mockCalc_ValueFromRangeWeighted(iNumber, CarData.Range("CarColors[Color]"), CarData.Range("CarColors[Popularity]"))
End Function

'Mock a value present in some range.
'@param iNumber - The number of items to generate
'@param values - The range to obtain a random value from
'@returns - A column of random items from the range
'@remark Useful for generating ID's from an existing range e.g. mocking relationships
Public Function mockCalc_ValueFromRange(ByVal iNumber As Long, values As Range) As Variant
  mockCalc_ValueFromRange = genColumn(iNumber, stdCallback.CreateFromPointer(AddressOf getRandomValueFromRange, vbVariant, Array(vbObject)).Bind(values))
End Function

'Mock a value present in some range, weighted by another range
'@param iNumber - The number of items to generate
'@param values - The range to obtain a random value from
'@param weights - The range to obtain the weights from
'@returns - A column of random items from the values range weighted by the weights range
'@remark Useful for generating ID's from an existing range e.g. mocking relationships especially many-to-many relationships
Public Function mockCalc_ValueFromRangeWeighted(ByVal iNumber As Long, values As Range, weights As Range) As Variant
  mockCalc_ValueFromRangeWeighted = genColumn(iNumber, stdCallback.CreateFromPointer(AddressOf getRandomValueFromRangeWeighted, vbVariant, Array(vbObject, vbObject)).Bind(values, weights))
End Function

'Obtain a random item from an array of weighted items
'@param iNumber - The number of items to generate
'@param vArray - An array of items and their weights.  The first item is the item, the second is the weight. [item1, weight1, item2, weight2, ...]
'@returns - A column of random items from the array
'@example mockCalc_WeightedArray(10, Array("A", 1, "B", 2, "C", 3))
Public Function mockCalc_WeightedArray(ByVal iNumber As Long, ParamArray vArray()) As Variant
  Dim vRet(): ReDim vRet(1 To iNumber, 1 To 1)
  Dim vArr: vArr = vArray
  For i = 1 To iNumber
    vRet(i, 1) = getRandomWeightedArrayItem(vArr)
  Next
  mockCalc_WeightedArray = vRet
End Function

'Mock a value compliant with a regex pattern
'@param iNumber - The number of items to generate
'@param sRegex - The regex pattern to generate values from
'@returns - A column of random items from the array
Public Function mockCalc_Regex(ByVal iNumber As Long, ByVal sRegex As String) as Variant
  Dim rx As stdRegex2: Set rx = stdRegex2.Create(sRegex, "")
  mockCalc_Regex = genColumn(iNumber, stdCallback.CreateFromObjectMethod(rx, "Generate"))
End Function

'Generate a random set of perlin noise values for a given set of X and Y coordinates
'@param Xs - The X coordinates
'@param Ys - The Y coordinates
'@param iMax - The maximum value
'@param iMin - The minimum value
'@param Seed - The seed to use for the random number generator
'@returns - A column of random perlin noise values
'@remark - Useful for generating a random topography / elevation maps
Public Function mockTopo_Elevation(ByVal Xs as Range, ByVal Ys as Range, Optional ByVal iMax As Long = 100, Optional ByVal iMin As Long = 0, Optional ByVal Seed as Long = 0) As Variant
  Dim iNumber as Long: iNumber = Xs.Rows.CountLarge
  Dim vRet(): ReDim vRet(1 To iNumber, 1 To 1)
  Dim vXs: vXs = Xs.value
  Dim vYs: vYs = Ys.value
  Dim i As Long: For i = 1 To iNumber
    vRet(i, 1) = PerlinNoise2D(vXs(i,1), vYs(i,1), 0.5, 20, 2.5, 4, Seed, 0, 0)
    vRet(i, 1) = PerlinNoise2D(vXs(i,1), vYs(i,1), 0.5, 2, 0.5, 4, Seed, 0, 0)
    vRet(i, 1) = PerlinNoise2D(vXs(i,1), vYs(i,1), 0.5, 1, 0.25, 4, Seed, 0, 0)
    vRet(i, 1) = PerlinNoise2D(vXs(i,1), vYs(i,1), 0.5, 0.25, 0.1, 4, Seed, 0, 0)
    vRet(i, 1) = vRet(i, 1)/3.35 * (iMax - iMin) + iMin
  Next
  mockTopo_Elevation = vRet
End Function

'Generate a perlin noise value for a given set of X and Y coordinates
'@param X - The X coordinate
'@param Y - The Y coordinate
'@param Persistence - The persistence of the noise
'@param Frequency - The frequency of the noise
'@param Amplitude - The amplitude of the noise
'@param Octaves - The number of octaves
'@param RandomSeed - The seed to use for the random number generator
'@param OffsetX - The X offset
'@param OffsetY - The Y offset
'@returns - A perlin noise value
Private Function PerlinNoise2D(ByVal X As Double, ByVal Y As Double, ByVal Persistence As Double, ByVal Frequency As Double, ByVal Amplitude As Double, ByVal Octaves As Long, ByVal RandomSeed As Long, ByVal OffsetX As Double, ByVal OffsetY As Double) As Double
  Dim Seed As Long: Seed = RandomSeed
  Dim n As Long
  For n = 0 To Octaves - 1
    Dim Frequency2 As Double: Frequency2 = Frequency ^ n
    Dim Amplitude2 As Double: Amplitude2 = Amplitude ^ n
    Dim X2 As Double: X2 = X * Frequency2 + OffsetX
    Dim Y2 As Double: Y2 = Y * Frequency2 + OffsetY
    Dim i As Long
    i = (Int(X2) + Int(Y2) * 57) + Seed
    i = (i Xor 13) * i
    Dim P As Double: P = (1.0# - ((i * (i * i * 15731 + 789221) + 1376312589) And &H7FFFFFFF) / 1073741824.0#)
    Dim Total As Double: Total = Total + P * Amplitude2
  Next
  PerlinNoise2D = Total * Persistence
End Function



'TODO: Consider file path - https://regex101.com/library/zWGLMP
'TODO: Consider youtube link - https://regex101.com/library/OY96XI


































'Generate a static column of data
'@param iRowCount - The number of rows to generate
'@param item - The item to place in every row
'@returns - The column of data
Private Function genStatic(ByVal iRowCount as Long, ByVal item as variant) as variant
  dim v(): redim v(1 to iRowCount, 1 to 1)
  dim i as long: for i = 1 to iRowCount
    v(i, 1) = item
  next
  genStatic = v
End Function

'Generate a column of data using a callback
'@param iRowCount - The number of rows to generate
'@param callback - A callback function that returns a value
'@returns - A column of data
'@example genColumn(10, stdCallback.CreateFromPointer(AddressOf getGUID, vbString))
'@remarks - This is useful for generating a column of random data using a callback which returns a random value. The callback is called once per row.
Private Function genColumn(ByVal iRowCount As Long, ByVal callback As stdICallable) As Variant
  Dim v(): ReDim v(1 To iRowCount, 1 To 1)
  For i = 1 To iRowCount
    v(i, 1) = callback.Run()
  Next
  genColumn = v
End Function

'Generate a random date between min and max
'@param iMaxDate - The maximum date
'@param iMinDate - The minimum date
'@returns - A random date between min and max
Private Function genDate(ByVal iMaxDate As Date, ByVal iMinDate As Date) As Date
  genDate = CDate(RandBetween(iMinDate, iMaxDate))
End Function

'Generate a random date between min and max, skewed with data quality issues as if it were written by a human
'@param iMaxDate - The maximum date
'@param iMinDate - The minimum date
'@returns - A random date between min and max, skewed with data quality issues as if it were written by a human
Private Function genDateSkewed(ByVal iMaxDate As Date, ByVal iMinDate As Date) As String
  static formats as variant, delims as variant, fub as long, flb as long, dub as long, dlb as long
  if isEmpty(formats) then 
    formats = Array("yyyy-mm-dd", "yyyy-mmm-dd", "yyyy-mmmm-dd", "dd-mm-yyyy", "dd-mmm-yyyy", "dd-mmmm-yyyy", "dd-mm-yy", "dd-mmm-yy", "dd-mmmm-yy", "dd-mmmyy")
    fub = ubound(formats)
    flb = lbound(formats)
    delims = Array("-", ".", " ", "/", "\", "_","")
    dub = ubound(delims)
    dlb = lbound(delims)
  end if
  Dim sFormat as string: sFormat = replace(formats(RandBetween(flb, fub)), "-", delims(RandBetween(dlb, dub)))
  genDateSkewed = format(genDate(iMaxDate, iMinDate), sFormat)
End Function

'Get a random value from a given range
'@param values - The range to obtain a random value from
'@returns - A random value from the range
Private Function getRandomValueFromRange(ByVal values As Range) As Variant
  Static cache As Object: If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
  If Not cache.Exists(values.Address) Then
    Set cache(values.Address) = CreateObject("Scripting.Dictionary")
    cache(values.Address)("values") = values.value
    cache(values.Address)("count") = values.Rows.CountLarge
  End If
  With cache(values.Address)
    getRandomValueFromRange = .Item("values")(RandBetween(1, .Item("count")), 1)
  End With
End Function

'Get a random value from a given range, weighted by another range
'@param values - The range to obtain a random value from
'@param weights - The range to obtain the weights from
'@returns - A random value from the range weighted by the weights
Private Function getRandomValueFromRangeWeighted(ByVal rValues As Range, ByVal rWeights As Range) As Variant
  Static cache As Object: If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
  If Not cache.Exists(rValues.Address) Then
    Set cache(rValues.Address) = CreateObject("Scripting.Dictionary")
    cache(rValues.Address)("values") = rValues.value
    
    'Calculate cumulative weights
    Dim weights As Variant: weights = rWeights.value
    Dim iWeightCount As Long: iWeightCount = rWeights.Rows.CountLarge
    Dim vSumWeights(): ReDim vSumWeights(1 To iWeightCount, 1 To 1)
    Dim i as long: For i = 1 to iWeightCount
      vSumWeights(i, 1) = vSumWeights(iif(i=1,1,i - 1), 1) + weights(i, 1)
    Next
    
    'Normalize weights
    For i = 1 to iWeightCount
      vSumWeights(i, 1) = vSumWeights(i, 1) / vSumWeights(iWeightCount, 1)
    next
    cache(rValues.Address)("weights") = vSumWeights

    cache(rValues.Address)("count") = rValues.Rows.CountLarge
  End If
  With cache(rValues.Address)
    Dim vWeights as variant: vWeights = .Item("weights")
    dim rand as double: rand = rnd()
    For i = 1 to .item("count")
      If rand <= vWeights(i, 1) Then
        getRandomValueFromRangeWeighted = .Item("values")(i, 1)
        Exit Function
      End If
    Next
  End With
End Function

'Obtain a random item from an array of weighted items
'@param weightedItems - An array of items and their weights.  The first item is the item, the second is the weight. [item1, weight1, item2, weight2, ...]
'@returns - A random item from the array
Private Function getRandomWeightedArrayItem(ByRef weightedItems) As Variant
  Randomize
  Dim rand As Double: rand = Rnd()
  For i = 0 To UBound(weightedItems) Step 2
    sum = sum + weightedItems(i + 1)
    If rand <= sum Then
      getRandomWeightedArrayItem = weightedItems(i)
      Exit Function
    End If
  Next
End Function

'Generate a random GUID
'@returns - A random GUID
Private Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function

'Generate a random number between min and max
'@param min - The minimum value
'@param max - The maximum value
'@returns - A random number between min and max
Private Function RandBetween(ByVal min As Long, ByVal max As Long) As Long
    Randomize
    RandBetween = Int((max - min + 1) * Rnd + min)
End Function

'Convert a string to proper case
'@param s - The string to convert
'@returns - The string in proper case
Private Function ProperCase(ByVal s As String) As String
  Dim v As Variant: v = Split(s, " ")
  Dim i As Long: For i = LBound(v) To UBound(v)
    v(i) = UCase$(Left$(v(i), 1)) & LCase$(Mid$(v(i), 2))
  Next
  ProperCase = Join(v, " ")
End Function



