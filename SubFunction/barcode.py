from SubFunction.vb2py.vbfunctions import *
from SubFunction.vb2py.vbdebug import *

"""Symbology set Arrays"""
global I
I = 1
j = Integer()
DataToPrint = String()
OnlyCorrectData = String()
Encoding = String()
WeightedTotal = Long()
WeightValue = Integer()
CurrentValue = Long()
CheckDigitValue = Integer()
Factor = Integer()
CheckDigit = Integer()
CurrentEncoding = String()
NewLine = String()
CurrentChar = String()
CurrentCharNum = Integer()
C128_StartA = String()
C128_StartB = String()
C128_StartC = String()
C128_Stop = String()
C128Start = String()
C128CheckDigit = String()
StartCode = String()
StopCode = String()
LeadingDigit = Integer()
EAN2AddOn = String()
EAN5AddOn = String()
EANAddOnToPrint = String()
HumanReadableText = String()
StringLength = Integer()
StringLength = 0
Char1 = String()
Char2 = String()
nwPattern = String()
CorrectFNC = Integer()
Idx = Integer()
PrintableString = String()
SetC128 = vbObjectInitialize(objtype=Variant)
SetITF = vbObjectInitialize(objtype=Variant)
Set39 = vbObjectInitialize(objtype=Variant)
SetCODABAR = vbObjectInitialize(objtype=Variant)
SetMSI = vbObjectInitialize(objtype=Variant)
SetPostnet = vbObjectInitialize(objtype=Variant)
SetPlanet = vbObjectInitialize(objtype=Variant)
Set93 = vbObjectInitialize(objtype=Variant)

def MODU10(M10NumberData):
    M10StringLength = Integer()

    M10OnlyCorrectData = String()

    M10Factor = Integer()

    M10WeightedTotal = Integer()

    M10CheckDigit = Integer()

    M10I = Integer()
    # This is a general MOD10 function like the one required
    # for UCC, EAN and UPC barcodes
    #********************************************************
    M10OnlyCorrectData = ''
    M10StringLength = Len(M10NumberData)
    #Check to make sure data is numeric and remove dashes, etc.
    for M10I in vbForRange(1, M10StringLength):
        #Add all numbers to OnlyCorrectData string
        #2006.2 BDA modified the next 2 lines for compatibility with different office versions
        CurrentCharNum = AscW(Mid(M10NumberData, M10I, 1))
        if CurrentCharNum > 47 and CurrentCharNum < 58:
            M10OnlyCorrectData = M10OnlyCorrectData + Mid(M10NumberData, M10I, 1)
    #Generate MOD 10 check digit
    M10Factor = 3
    M10WeightedTotal = 0
    M10StringLength = Len(M10NumberData)
    for M10I in vbForRange(M10StringLength, 1, - 1):
        #Get the value of each number starting at the end
        #CurrentCharNum = Mid(M10NumberData, I, 1)
        #Multiply by the weighting factor which is 3,1,3,1...
        #and add the sum together
        M10WeightedTotal = M10WeightedTotal +  ( Val(Mid(M10NumberData, M10I, 1)) * M10Factor )
        #Change factor for next calculation
        M10Factor = 4 - M10Factor
    #Find the CheckDigit by finding the smallest number that = a multiple of 10
    M10I = ( M10WeightedTotal % 10 )
    if M10I != 0:
        M10CheckDigit = ( 10 - M10I )
    else:
        M10CheckDigit = 0
    fn_return_value = Str(M10CheckDigit)
    return fn_return_value

def IDAutomation_Uni_ProcessTilde(StringToProcess):
    OutString = String()

    CharsAdded = Integer()
    fn_return_value = ''
    StringLength = Len(StringToProcess)
    for I in vbForRange(1, StringLength):
        if ( I < StringLength - 2 )  and Mid(StringToProcess, I, 2) == '~m' and IsNumeric(Mid(StringToProcess, I + 2, 2)):
            WeightValue = Val(Mid(StringToProcess, I + 2, 2))
            #Dim CharsAdded As Integer
            for j in vbForRange(I, 1, - 1):
                if IsNumeric(Mid(OutString, j, 1)):
                    StringToCheck = StringToCheck + Mid(OutString, j, 1)
                    CharsAdded = CharsAdded + 1
                #when the number of digits added to StringToCheck equals the weight value exit the for loop
                if CharsAdded == WeightValue:
                    break
            CheckDigitValue = MODU10(StrReverse(StringToCheck))
            OutString = OutString + ChrW(CheckDigitValue + 48)
            I = I + 3
        elif  ( I < StringLength - 2 )  and Mid(StringToProcess, I, 1) == '~' and IsNumeric(Mid(StringToProcess, I + 1, 3)):
            CurrentCharNum = Val(Mid(StringToProcess, I + 1, 3))
            OutString = OutString + ChrW(CurrentCharNum)
            I = I + 3
            #This ElseIf was modified to support using () to add in AIs
        elif  ( I < StringLength - 4 )  and Mid(StringToProcess, I, 1) == '(' and  ( Mid(StringToProcess, I + 2, 1) == ')' or Mid(StringToProcess, I + 3, 1) == ')' or Mid(StringToProcess, I + 4, 1) == ')' or  ( Mid(StringToProcess, I + 5, 1) == ')' and  ( I < StringLength - 5 ) )  or  ( Mid(StringToProcess, I + 6, 1) == ')' and  ( I < StringLength - 6 ) )  or  ( Mid(StringToProcess, I + 7, 1) == ')' and  ( I < StringLength - 4 ) )  or  ( Mid(StringToProcess, I + 8, 1) == ')' and  ( I < StringLength - 4 ) ) ):
            #Assign ASCII 212-217 depending on how many digits between ()
            if Mid(StringToProcess, I + 3, 1) == ')':
                OutString = OutString + ChrW(212)
            if Mid(StringToProcess, I + 4, 1) == ')':
                OutString = OutString + ChrW(213)
            if Mid(StringToProcess, I + 5, 1) == ')':
                OutString = OutString + ChrW(214)
            if Mid(StringToProcess, I + 6, 1) == ')':
                OutString = OutString + ChrW(215)
            if Mid(StringToProcess, I + 7, 1) == ')':
                OutString = OutString + ChrW(216)
            if Mid(StringToProcess, I + 8, 1) == ')':
                OutString = OutString + ChrW(217)
            #This ElseIf was modified to exclude ")" from being encoded
        elif  ( I < StringLength - 2 )  and Mid(StringToProcess, I, 1) == ')':
            #Skip this character by breaking out of the else if
            pass
        else:
            OutString = OutString + Mid(StringToProcess, I, 1)
    fn_return_value = OutString
    StringToProcess = ''
    return fn_return_value

def IDAutomation_Uni_C128(DataToFormat, ApplyTilde=False):
    global I
    StringLength = 0
    DataToEncode = String()
    PrintableString = String()
    CorrectFNC = 0
    #in case ApplyTilde is null, set it to false
    if ApplyTilde != True:
        ApplyTilde = False
    PrintableString = ''
    DataToEncode = DataToFormat
    #2007.8 TB added the next line to move code to the IDAutomation_Uni_ProcessTilde function
    if ApplyTilde:
        DataToEncode = IDAutomation_Uni_ProcessTilde(DataToEncode)
    DataToFormat = DataToEncode
    DataToEncode = ''
    SetC128 = Array('EFF', 'FEF', 'FFE', 'BBG', 'BCF', 'CBF', 'BFC', 'BGB', 'CFB', 'FBC', 'FCB', 'GBB', 'AFJ', 'BEJ', 'BFI', 'AJF',
                    'BIF', 'BJE', 'FJA', 'FAJ', 'FBI', 'EJB', 'FIB', 'IEI', 'IBF', 'JAF', 'JBE', 'IFB', 'JEB', 'JFA', 'EEG', 'EGE',
                    'GEE', 'ACG', 'CAG', 'CCE', 'AGC', 'CEC', 'CGA', 'ECC', 'GAC', 'GCA', 'AEK', 'AGI', 'CEI', 'AIG', 'AKE', 'CIE',
                    'IIE', 'ECI', 'GAI', 'EIC', 'EKA', 'EII', 'IAG', 'ICE', 'KAE', 'IEC', 'IGA', 'KEA', 'IMA', 'FDA', 'OAA', 'ABH',
                    'ADF', 'BAH', 'BDE', 'DAF', 'DBE', 'AFD', 'AHB', 'BED', 'BHA', 'DEB', 'DFA', 'HBA', 'FAD', 'MIA', 'HAB', 'CMA',
                    'ABN', 'BAN', 'BBM', 'ANB', 'BMB', 'BNA', 'MBB', 'NAB', 'NBA', 'EEM', 'EME', 'MEE', 'AAO', 'ACM', 'CAM', 'AMC',
                    'AOA', 'MAC', 'MCA', 'AIM', 'AMI', 'IAM', 'MAI', 'EDB', 'EBD', 'EBJ')
    #Here we select character set A, B or C for the START character
    #Start A = "EDB"
    #Start B = "EBD"
    #Start C = "EBJ"
    StringLength = Len(DataToFormat)
    CurrentCharNum = AscW(Mid(DataToFormat, 1, 1))
    #Set A?
    if CurrentCharNum < 32:
        C128Start = 'EDB'
    #Set B?
    if CurrentCharNum > 31 and CurrentCharNum < 127:
        C128Start = 'EBD'
    if CurrentCharNum == 197:
        C128Start = 'EBD'
    #Updated V5.08 by BDA
    #Set C?
    if ( ( StringLength > 3 )  and IsNumeric(Mid(DataToFormat, 1, 4)) ):
        C128Start = 'EBJ'
    #202 & 212-215 is for the FNC1, with this Start C is mandatory
    if CurrentCharNum == 202:
        C128Start = 'EBJ'
    if CurrentCharNum > 211:
        C128Start = 'EBJ'
    if C128Start == 'EDB':
        CurrentEncoding = 'A'
    if C128Start == 'EBD':
        CurrentEncoding = 'B'
    if C128Start == 'EBJ':
        CurrentEncoding = 'C'
    I = 1
    while I <= StringLength:
        CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
        #check for FNC1 in any set which is ASCII 202 and ASCII 212-215
        if ( ( CurrentCharNum > 201 )  and  ( CurrentCharNum < 219 ) ):
            DataToEncode = DataToEncode + ChrW(202)
            #check for switching to character set C
        elif CurrentCharNum == 195:
            if CurrentEncoding == 'C':
                DataToEncode = DataToEncode + ChrW(200)
                CurrentEncoding = 'B'
            DataToEncode = DataToEncode + ChrW(195)
        elif CurrentCharNum == 196:
            if CurrentEncoding == 'C':
                DataToEncode = DataToEncode + ChrW(200)
                CurrentEncoding = 'B'
            DataToEncode = DataToEncode + ChrW(196)
        elif CurrentCharNum == 197:
            if CurrentEncoding == 'C':
                DataToEncode = DataToEncode + ChrW(200)
                CurrentEncoding = 'B'
            DataToEncode = DataToEncode + ChrW(197)
        elif CurrentCharNum == 198:
            if CurrentEncoding == 'C':
                DataToEncode = DataToEncode + ChrW(200)
                CurrentEncoding = 'B'
            DataToEncode = DataToEncode + ChrW(198)
        elif CurrentCharNum == 200:
            if CurrentEncoding == 'C':
                DataToEncode = DataToEncode + ChrW(200)
                CurrentEncoding = 'B'
            DataToEncode = DataToEncode + ChrW(200)
        elif  (((I < StringLength - 2)  and  (IsNumeric(Mid(DataToFormat, I, 1)))  and  (IsNumeric(Mid(DataToFormat, I + 1, 1)))  and  (IsNumeric(Mid(DataToFormat, I, 4))) )  or
                    ((I < StringLength)  and  (IsNumeric(Mid(DataToFormat, I, 1)))  and  (IsNumeric(Mid(DataToFormat, I + 1, 1)))  and  (CurrentEncoding == 'C'))):
            #Updated V5.12 by BDA 11-22-2005
            #check to see if we have an odd number of numbers to encode,
            #if so stay in current set and then switch to save space
            if CurrentEncoding != 'C':
                j = I
                Factor = 3
                while j <= StringLength and IsNumeric(Mid(DataToFormat, j, 1)):
                    Factor = 4 - Factor
                    j = j + 1
                if Factor == 1:
                    #if so stay in current set for 1 character to save space
                    DataToEncode = DataToEncode + ChrW(CurrentCharNum)
                    I = I + 1
            #switch to set C if not already in it
            if CurrentEncoding != 'C':
                DataToEncode = DataToEncode + ChrW(199)
            CurrentEncoding = 'C'
            CurrentChar = Mid(DataToFormat, I, 2)
            CurrentValue = int(Val(CurrentChar))
            #set the CurrentValue to the number of String CurrentChar
            DataToEncode = DataToEncode + ChrW(CurrentValue + 32)
            I = I + 2
            if I > StringLength:
                break
            #check for switching to character set A
        elif  (I <= StringLength)  and  ((AscW(Mid(DataToFormat, I, 1)) < 31)  or  
                ((CurrentEncoding == 'A')  and  (AscW(Mid(DataToFormat, I, 1)) > 32 and  (AscW(Mid(DataToFormat, I, 1)))  < 96 ))):
            #switch to set A if not already in it
            if CurrentEncoding != 'A':
                DataToEncode = DataToEncode + ChrW(201)
            CurrentEncoding = 'A'
            #Get the ASCII value of the next character
            CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
            if CurrentCharNum < 32:
                DataToEncode = DataToEncode + ChrW(CurrentCharNum + 96)
            elif CurrentCharNum > 32:
                DataToEncode = DataToEncode + ChrW(CurrentCharNum)
            I = I + 1
            if I > StringLength:
                break
            #check for switching to character set B
        elif  (I <= StringLength)  and  (((AscW(Mid(DataToFormat, I, 1)))  > 31)  and  (( AscW(Mid(DataToFormat, I, 1))))  < 127):
            #switch to set B if not already in it
            if CurrentEncoding != 'B':
                DataToEncode = DataToEncode + ChrW(200)
            CurrentEncoding = 'B'
            #Get the ASCII value of the next character
            CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
            DataToEncode = DataToEncode + ChrW(CurrentCharNum)
            I = I + 1
            if I > StringLength:
                break
    #<<<< Calculate Modulo 103 Check Digit >>>>
    if C128Start == 'EDB':
        WeightedTotal = 103
        #CurrentEncoding = "A"
    if C128Start == 'EBD':
        WeightedTotal = 104
        #CurrentEncoding = "B"
    if C128Start == 'EBJ':
        WeightedTotal = 105
        #CurrentEncoding = "C"
    StringLength = Len(DataToEncode)
    for I in vbForRange(1, StringLength):
        CurrentCharNum = AscW(Mid(DataToEncode, I, 1))
        if CurrentCharNum < 135:
            CurrentValue = CurrentCharNum - 32
        if CurrentCharNum > 134:
            CurrentValue = CurrentCharNum - 100
        if CurrentCharNum == 194:
            CurrentValue = 0
        PrintableString = PrintableString + SetC128(CurrentValue)
        CurrentValue = CurrentValue * I
        WeightedTotal = WeightedTotal + CurrentValue
    # print(WeightedTotal% 103)
    CheckDigitValue = ( WeightedTotal % 103 )
    # print(SetC128(CheckDigitValue))
    DataToEncode = ''
    #GIAH produces the stop character.
    #Formula = C128Start & PrintableString & SetC128(CheckDigitValue) & ChrW(206) & "GIAH" & "j"
    fn_return_value = C128Start + PrintableString + SetC128(CheckDigitValue) + 'GIAH'
    PrintableString = ''
    return fn_return_value

def IDAutomation_Uni_C39(DataToEncode, N_Dimension=2, IncludeCheckDigit=False):
    CurrentEncoding = ''
    DataToPrint = ''
    OnlyCorrectData = ''
    if N_Dimension != 3 and N_Dimension != 2.5:
        N_Dimension = 2
    Set39 = Array('nnnwwnwnnn', 'wnnwnnnnwn', 'nnwwnnnnwn', 'wnwwnnnnnn', 'nnnwwnnnwn',
                    'wnnwwnnnnn', 'nnwwwnnnnn', 'nnnwnnwnwn', 'wnnwnnwnnn', 'nnwwnnwnnn',
                    'wnnnnwnnwn', 'nnwnnwnnwn', 'wnwnnwnnnn', 'nnnnwwnnwn', 'wnnnwwnnnn',
                    'nnwnwwnnnn', 'nnnnnwwnwn', 'wnnnnwwnnn', 'nnwnnwwnnn', 'nnnnwwwnnn',
                    'wnnnnnnwwn', 'nnwnnnnwwn', 'wnwnnnnwnn', 'nnnnwnnwwn', 'wnnnwnnwnn',
                    'nnwnwnnwnn', 'nnnnnnwwwn', 'wnnnnnwwnn', 'nnwnnnwwnn', 'nnnnwnwwnn',
                    'wwnnnnnnwn', 'nwwnnnnnwn', 'wwwnnnnnnn', 'nwnnwnnnwn', 'wwnnwnnnnn',
                    'nwwnwnnnnn', 'nwnnnnwnwn', 'wwnnnnwnnn', 'nwwnnnwnnn', 'nwnwnwnnnn',
                    'nwnwnnnwnn', 'nwnnnwnwnn', 'nnnwnwnwnn', 'nwnnwnwnnn')
    DataToEncode = UCase(DataToEncode)
    StringLength = Len(DataToEncode)
    WeightedTotal = 0
    for I in vbForRange(1, StringLength):
        #Get the value of each number
        CurrentCharNum = ( AscW(Mid(DataToEncode, I, 1)) )
        CurrentEncoding = ''
        #Set CurrentValue to an invalid value by default
        CurrentValue = 99
        #Get the C39 value of CurrentChar
        #0-9
        if CurrentCharNum < 58 and CurrentCharNum > 47:
            CurrentValue = CurrentCharNum - 48
        #A-Z
        if CurrentCharNum < 91 and CurrentCharNum > 64:
            CurrentValue = CurrentCharNum - 55
        #Space
        if CurrentCharNum == 32:
            CurrentValue = 38
        #-
        if CurrentCharNum == 45:
            CurrentValue = 36
        #.
        if CurrentCharNum == 46:
            CurrentValue = 37
        #$
        if CurrentCharNum == 36:
            CurrentValue = 39
        #/
        if CurrentCharNum == 47:
            CurrentValue = 40
        #+
        if CurrentCharNum == 43:
            CurrentValue = 41
        #%
        if CurrentCharNum == 37:
            CurrentValue = 42
        #add the values together for the check digit
        if CurrentValue != '99' and IncludeCheckDigit:
            WeightedTotal = WeightedTotal + CurrentValue
        #retrieve the pattern, for example, 0 = "nnnwwnwnnn"
        if CurrentValue != '99':
            nwPattern = Set39(CurrentValue)
        if N_Dimension == 3 and CurrentValue != '99':
            for j in vbForRange(1, 10, 2):
                CurrentEncoding = Mid(nwPattern, j, 2)
                select_variable_0 = CurrentEncoding
                if (select_variable_0 == 'nn'):
                    DataToPrint = DataToPrint + 'A'
                elif (select_variable_0 == 'nw'):
                    DataToPrint = DataToPrint + 'C'
                elif (select_variable_0 == 'wn'):
                    DataToPrint = DataToPrint + 'I'
                elif (select_variable_0 == 'ww'):
                    DataToPrint = DataToPrint + 'K'
        if N_Dimension == 2.5 and CurrentValue != '99':
            for j in vbForRange(1, 10, 2):
                CurrentEncoding = Mid(nwPattern, j, 2)
                select_variable_1 = CurrentEncoding
                if (select_variable_1 == 'nn'):
                    DataToPrint = DataToPrint + 'W'
                elif (select_variable_1 == 'nw'):
                    DataToPrint = DataToPrint + 'X'
                elif (select_variable_1 == 'wn'):
                    DataToPrint = DataToPrint + 'Y'
                elif (select_variable_1 == 'ww'):
                    DataToPrint = DataToPrint + 'Z'
        if N_Dimension == 2 and CurrentValue != '99':
            for j in vbForRange(1, 10, 2):
                CurrentEncoding = Mid(nwPattern, j, 2)
                select_variable_2 = CurrentEncoding
                if (select_variable_2 == 'nn'):
                    DataToPrint = DataToPrint + 'A'
                elif (select_variable_2 == 'nw'):
                    DataToPrint = DataToPrint + 'B'
                elif (select_variable_2 == 'wn'):
                    DataToPrint = DataToPrint + 'E'
                elif (select_variable_2 == 'ww'):
                    DataToPrint = DataToPrint + 'F'
    if IncludeCheckDigit:
        #divide the WeightedTotal by 43 and get the remainder, this is the CheckDigit
        CheckDigitValue = ( WeightedTotal % 43 )
        nwPattern = Set39(CheckDigitValue)
        if N_Dimension == 3 and CurrentValue != '99':
            for j in vbForRange(1, 10, 2):
                CurrentEncoding = Mid(nwPattern, j, 2)
                select_variable_3 = CurrentEncoding
                if (select_variable_3 == 'nn'):
                    DataToPrint = DataToPrint + 'A'
                elif (select_variable_3 == 'nw'):
                    DataToPrint = DataToPrint + 'C'
                elif (select_variable_3 == 'wn'):
                    DataToPrint = DataToPrint + 'I'
                elif (select_variable_3 == 'ww'):
                    DataToPrint = DataToPrint + 'K'
        if N_Dimension == 2.5 and CurrentValue != '99':
            for j in vbForRange(1, 10, 2):
                CurrentEncoding = Mid(nwPattern, j, 2)
                select_variable_4 = CurrentEncoding
                if (select_variable_4 == 'nn'):
                    DataToPrint = DataToPrint + 'W'
                elif (select_variable_4 == 'nw'):
                    DataToPrint = DataToPrint + 'X'
                elif (select_variable_4 == 'wn'):
                    DataToPrint = DataToPrint + 'Y'
                elif (select_variable_4 == 'ww'):
                    DataToPrint = DataToPrint + 'Z'
        if N_Dimension == 2 and CurrentValue != '99':
            for j in vbForRange(1, 10, 2):
                CurrentEncoding = Mid(nwPattern, j, 2)
                select_variable_5 = CurrentEncoding
                if (select_variable_5 == 'nn'):
                    DataToPrint = DataToPrint + 'A'
                elif (select_variable_5 == 'nw'):
                    DataToPrint = DataToPrint + 'B'
                elif (select_variable_5 == 'wn'):
                    DataToPrint = DataToPrint + 'E'
                elif (select_variable_5 == 'ww'):
                    DataToPrint = DataToPrint + 'F'
    #Add start and stop bars
    if N_Dimension == 3:
        fn_return_value = 'CAIIA' + DataToPrint + 'CAIIA'
    if N_Dimension == 2.5:
        fn_return_value = 'XWYYW' + DataToPrint + 'XWYYW'
    if N_Dimension == 2:
        fn_return_value = 'BAEEA' + DataToPrint + 'BAEEA'
    return fn_return_value
