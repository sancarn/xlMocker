Attribute VB_Name = "mockSentence"


'Generates a sentence recursively. A sentence is defined as:
'   <sentence> = <regular-sentence> | <question> | <command> | <exclamatory-sentence>
'   <subject-phrase> = <noun-phrase> | <pronoun>
'   <noun-phrase> = <complex-noun> | <complex-noun> <preposition-phrase> | <noun-phrase> <conjunction> <noun-phrase>
'   <verb-phrase> = <complex-verb> | <complex-verb> <object-phrase>
'   <object-phrase> = <noun-phrase> | <pronoun>
'   <preposition-phrase> = <preposition> <noun-phrase>
'   <complex-noun> = (<article> <adjective-phrase>? <noun> | <adjective-phrase>? <noun> | <compound-noun>) <preposition-phrase>?
'   <complex-verb> = <adverb-phrase>? <verb> | <adverb-phrase>? <verb> <object-phrase> | <verb> <adverb-phrase>
'   <adjective-phrase> = <adjective> | <adjective> <adjective-phrase>
'   <adverb-phrase> = <adverb> | <adverb> <adverb-phrase>
'   <article> = a | an | the | <quantifier>
'   <quantifier> = some | any | every | all | no | <number>
'   <number> = <number>? (1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9)
'   <pronoun> = I | you | he | she | it | we | they | me | him | her | us | them | this | that | these | those | ...
'   <conjunction> = and | or | but | nor | yet | so
'   <compound-noun> = <noun> <noun> | <noun> <compound-noun>
'   <question> = <question-word> <subject-auxiliary-inversion> <verb-phrase>? "?"
'   <question-word> = who | what | where | when | why | how | which | whom | whose
'   <auxiliary-verb> = is | am | are | was | were | do | does | did | have | has | had | can | could | will | would | shall | should | must | might | ought
'   <subject-auxiliary-inversion> = <auxiliary-verb> <subject-phrase> | <subject-phrase> <auxiliary-verb>
'   <command> = "Please" <verb-phrase>
'   <exclamatory-sentence> = <subject-phrase> <verb-phrase> "!"
'   <regular-sentence> = <subject-phrase> <verb-phrase> "."


Public Function GenerateParagraph() As String
    Dim paragraph As String
    Dim i As Long: For i = 1 To RandBetween(2, 10)
        paragraph = paragraph & " " & GenerateSentence()
    Next
    
    GenerateParagraph = trim(paragraph)
End Function

Private Function GenerateSentence() As String
    select case RandBetween(1, 4)
        case 1: GenerateSentence = GenerateRegularSentence()
        case 2: GenerateSentence = GenerateQuestion()
        case 3: GenerateSentence = GenerateCommand()
        case 4: GenerateSentence = GenerateExclamatorySentence()
    end select
End Function

Private Function GenerateRegularSentence() As String
    GenerateRegularSentence = GenerateSubjectPhrase() & " " & GenerateVerbPhrase() & "."
End Function

Private Function GenerateQuestion() As String
    GenerateQuestion = GenerateQuestionWord() & " " & GenerateSubjectAuxiliaryInversion() & " " & GenerateVerbPhrase() & "?"
End Function

Private Function GenerateCommand() As String
    GenerateCommand = "Please " & GenerateVerbPhrase() & "."
End Function

Private Function GenerateExclamatorySentence() As String
    GenerateExclamatorySentence = GenerateSubjectPhrase() & " " & GenerateVerbPhrase() & "!"
End Function

Private Function GenerateSubjectPhrase() As String
    select case RandBetween(1, 2)
        case 1: GenerateSubjectPhrase = GenerateNounPhrase()
        case 2: GenerateSubjectPhrase = GeneratePronoun()
    end select
End Function

Private Function GenerateNounPhrase() As String
    select case RandBetween(1, 3)
        case 1: GenerateNounPhrase = GenerateComplexNoun()
        case 2: GenerateNounPhrase = GenerateComplexNoun() & " " & GeneratePrepositionPhrase()
        case 3: GenerateNounPhrase = GenerateNounPhrase() & " " & GenerateConjunction() & " " & GenerateNounPhrase()
    end select
End Function

Private Function GenerateVerbPhrase() As String
    select case RandBetween(1, 3)
        case 1: GenerateVerbPhrase = GenerateComplexVerb()
        case 2: GenerateVerbPhrase = GenerateComplexVerb() & " " & GenerateObjectPhrase()
        case 3: GenerateVerbPhrase = GenerateVerbPhrase() & " " & GenerateAdverbPhrase()
    end select
End Function

Private Function GenerateObjectPhrase() As String
    select case RandBetween(1, 2)
        case 1: GenerateObjectPhrase = GenerateNounPhrase()
        case 2: GenerateObjectPhrase = GeneratePronoun()
    end select
End Function

Private Function GeneratePrepositionPhrase() As String
    GeneratePrepositionPhrase = GeneratePreposition() & " " & GenerateNounPhrase()
End Function

Private Function GenerateComplexNoun() As String
    select case RandBetween(1, 2)
        case 1: GenerateComplexNoun = GenerateArticle() & " " & GenerateAdjectivePhrase() & " " & GenerateNoun()
        case 2: GenerateComplexNoun = GenerateAdjectivePhrase() & " " & GenerateNoun()
    end select
    
    if RandBetween(1, 2) = 1 then GenerateComplexNoun = GenerateComplexNoun & " " & GeneratePrepositionPhrase()
End Function

Private Function GenerateComplexVerb() As String
    select case RandBetween(1, 3)
        case 1: GenerateComplexVerb = GenerateAdverbPhrase() & " " & GenerateVerb()
        case 2: GenerateComplexVerb = GenerateAdverbPhrase() & " " & GenerateVerb() & " " & GenerateObjectPhrase()
        case 3: GenerateComplexVerb = GenerateVerb() & " " & GenerateAdverbPhrase()
    end select
End Function

Private Function GenerateAdjectivePhrase() As String
    select case RandBetween(1, 2)
        case 1: GenerateAdjectivePhrase = GenerateAdjective()
        case 2: GenerateAdjectivePhrase = GenerateAdjective() & " " & GenerateAdjectivePhrase()
    end select
End Function

Private Function GenerateAdverbPhrase() As String
    select case RandBetween(1, 2)
        case 1: GenerateAdverbPhrase = GenerateAdverb()
        case 2: GenerateAdverbPhrase = GenerateAdverb() & " " & GenerateAdverbPhrase()
    end select
End Function

Private Function GenerateArticle() As String
    select case RandBetween(1, 4)
        case 1: GenerateArticle = "a"
        case 2: GenerateArticle = "an"
        case 3: GenerateArticle = "the"
        case 4: GenerateArticle = GenerateQuantifier()
    end select
End Function

Private Function GenerateQuestionWord() As String
    select case RandBetween(1, 9)
        case 1: GenerateQuestionWord = "who"
        case 2: GenerateQuestionWord = "what"
        case 3: GenerateQuestionWord = "where"
        case 4: GenerateQuestionWord = "when"
        case 5: GenerateQuestionWord = "why"
        case 6: GenerateQuestionWord = "how"
        case 7: GenerateQuestionWord = "which"
        case 8: GenerateQuestionWord = "whom"
        case 9: GenerateQuestionWord = "whose"
    end select
End Function

Private Function GenerateSubjectAuxiliaryInversion() As String
    select case RandBetween(1, 2)
        case 1: GenerateSubjectAuxiliaryInversion = GenerateAuxiliaryVerb() & " " & GenerateSubjectPhrase()
        case 2: GenerateSubjectAuxiliaryInversion = GenerateSubjectPhrase() & " " & GenerateAuxiliaryVerb()
    end select
End Function

Private Function GeneratePronoun() As String
    static pronouns As Variant: if isEmpty(pronouns) then pronouns = Range("pronouns").Value
    pronouns = pronouns(RandBetween(1, ubound(pronouns)), 1)
End Function

Private Function GenerateAuxiliaryVerb() As String
    select case RandBetween(1, 20)
        case 1: GenerateAuxiliaryVerb = "is"
        case 2: GenerateAuxiliaryVerb = "am"
        case 3: GenerateAuxiliaryVerb = "are"
        case 4: GenerateAuxiliaryVerb = "was"
        case 5: GenerateAuxiliaryVerb = "were"
        case 6: GenerateAuxiliaryVerb = "do"
        case 7: GenerateAuxiliaryVerb = "does"
        case 8: GenerateAuxiliaryVerb = "did"
        case 9: GenerateAuxiliaryVerb = "have"
        case 10: GenerateAuxiliaryVerb = "has"
        case 11: GenerateAuxiliaryVerb = "had"
        case 12: GenerateAuxiliaryVerb = "can"
        case 13: GenerateAuxiliaryVerb = "could"
        case 14: GenerateAuxiliaryVerb = "will"
        case 15: GenerateAuxiliaryVerb = "would"
        case 16: GenerateAuxiliaryVerb = "shall"
        case 17: GenerateAuxiliaryVerb = "should"
        case 18: GenerateAuxiliaryVerb = "must"
        case 19: GenerateAuxiliaryVerb = "might"
        case 20: GenerateAuxiliaryVerb = "ought"
    end select
End Function

Private Function GenerateNoun() as String
    static nouns As Variant: if isEmpty(nouns) then nouns = Range("nouns").Value
    GenerateNoun = Nouns(RandBetween(1, ubound(Nouns)), 1)
end Function

Private Function GenerateVerb() as String
    static verbs As Variant: if isEmpty(verbs) then verbs = Range("verbs").Value
    GenerateVerb = Verbs(RandBetween(1, ubound(Verbs)), 1)
end Function
Private Function GenerateAdverb() as String
    static adverbs As Variant: if isEmpty(adverbs) then adverbs = Range("adverbs").Value
    GenerateAdverb = Adverbs(RandBetween(1, ubound(Adverbs)), 1)
end Function
Private Function GenerateAdjective() as String
    static adjectives As Variant: if isEmpty(adjectives) then adjectives = Range("adjectives").Value
    GenerateAdjective = Adjectives(RandBetween(1, ubound(Adjectives)), 1)
end Function
Private Function GeneratePreposition() as String
    static prepositions As Variant: if isEmpty(prepositions) then prepositions = Range("prepositions").Value
    GeneratePreposition = Prepositions(RandBetween(1, ubound(Prepositions)), 1)
end Function
Private Function GenerateConjunction() as String
    static conjunctions As Variant: if isEmpty(conjunctions) then conjunctions = Range("conjunctions").Value
    GenerateConjunction = Conjunctions(RandBetween(1, ubound(Conjunctions)), 1)
end Function


Private Function GenerateQuantifier() as String
    select case RandBetween(1, 6)
        case 1: GenerateQuantifier = "some"
        case 2: GenerateQuantifier = "any"
        case 3: GenerateQuantifier = "every"
        case 4: GenerateQuantifier = "all"
        case 5: GenerateQuantifier = "no"
        case 6: GenerateQuantifier = GenerateNumber()
    end select
end Function

Private Function GenerateNumber() as String
    GenerateNumber = RandBetween(1, 100)
end Function




'Generate a random number between min and max
'@param min - The minimum value
'@param max - The maximum value
'@returns - A random number between min and max
Private Function RandBetween(ByVal min As Long, ByVal max As Long) As Long
    RandBetween = Int((max - min + 1) * Rnd + min)
End Function