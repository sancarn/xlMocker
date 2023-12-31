Attribute VB_Name = "mockLoremIpsum"

Public Function GenerateLoremIpsum(Optional ByVal iMaxParagraphs as Long = 1, Optional ByVal iMinParagraphs as Long = 1, Optional ByVal iMaxSentences as long = 4, Optional ByVal iMinSentences as Long = 2, Optional ByVal iMaxWords as Long = 10, Optional ByVal iMinWords as Long = 5) as String
    Dim i as Long
    For i = 1 to RandBetween(iMinParagraphs, iMaxParagraphs)
        GenerateLoremIpsum = GenerateLoremIpsum & GenerateParagraph(iMinSentences, iMaxSentences, iMinWords, iMaxWords) & vbCrLf & vbCrLf
    Next
    GenerateLoremIpsum = Left(GenerateLoremIpsum, Len(GenerateLoremIpsum) - 4)
End Function

Private Function GenerateParagraph(ByVal iMinSentences as Long, ByVal iMaxSentences as long, ByVal iMinWords as Long, Optional ByVal iMaxWords as Long) as String
    Dim i as Long
    For i = 1 to RandBetween(iMinSentences, iMaxSentences)
        GenerateParagraph = GenerateParagraph & GenerateSentence(iMinWords, iMaxWords)
    Next
    GenerateParagraph = Left(GenerateParagraph, Len(GenerateParagraph) - 1)
End Function

Private Function GenerateSentence(ByVal iMinWords as Long, Optional ByVal iMaxWords as Long) as String
    Dim i as Long
    For i = 1 to RandBetween(iMinWords, iMaxWords)
        GenerateSentence = GenerateSentence & GenerateWord() & " "
    Next
    GenerateSentence = Left(GenerateSentence, Len(GenerateSentence) - 1) & ". "
    Mid(GenerateSentence, 1, 1) = UCase(Left(GenerateSentence, 1))
End Function

Private Function GenerateWord() as String
    static words as Variant: if isEmpty(words) then words = Array("ad","adipisicing","aliqua","aliquip","amet","anim","aute","cillum","commodo","consectetur","consequat","culpa","cupidatat","deserunt","do","dolor","dolore","duis","ea","eiusmod","elit","enim","esse","est","et","eu","ex","excepteur","exercitation","fugiat","id","in","incididunt","ipsum","irure","labore","laboris","laborum","lorem","magna","minim","mollit","nisi","non","nostrud","nulla","occaecat","officia","pariatur","proident","qui","quis","reprehenderit","sint","sit","sunt","tempor","ullamco","ut","velit","veniam", "voluptate")
    GenerateWord = words(RandBetween(LBound(words), UBound(words)))
End Function


Private Function RandBetween(ByVal iMin as Long, ByVal iMax as Long) as Long
    Randomize
    RandBetween = Int((iMax - iMin + 1) * Rnd + iMin)
End Function