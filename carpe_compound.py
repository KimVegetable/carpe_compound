# carpe_compound.py
import datetime

import compoundfiles
import os
import struct

class Compound:

    ### Dameged Documents ###
    CONST_DOCUMENT_NORMAL = 0x0000
    CONST_DOCUMENT_DAMAGED = 0x0001
    CONST_DOCUMENT_UNKNOWN_DAMAGED = 0x0002

    ### Encrypted Documents ###
    CONST_DOCUMENT_NO_ENCRYPTED = 0x0000
    CONST_DOCUMENT_ENCRYPTED = 0x0001
    CONST_DOCUMENT_UNKNOWN_ENCRYPTED = 0x0002

    ### Restoreable Documents ###
    CONST_DOCUMENT_RESTORABLE = 0x0000
    CONST_DOCUMENT_UNRESTORABLE = 0x0001
    CONST_DOCUMENT_UNKNOWN_RESTORABLE = 0x0002

    CONST_SUCCESS = True
    CONST_ERROR = False

    def __init__(self, filePath):
        if(os.path.exists(filePath)):
            try:
                self.fp = compoundfiles.CompoundFileReader(filePath)
                self.isDamaged = self.CONST_DOCUMENT_NORMAL
                print("Normal File exist!!")
            except compoundfiles.errors.CompoundFileInvalidBomError:
                self.fp = open(filePath, 'rb')
                self.isDamaged = self.CONST_DOCUMENT_DAMAGED
                print("Damaged File exist!!")

        else:
            self.fp = None
            print("File doesn't exist.")

        self.fileSize = os.path.getsize(filePath)
        self.fileName = os.path.basename(filePath)
        self.filePath = filePath
        self.fileType = os.path.splitext(filePath)[1][1:]   # delete '.' in '.xls' r
        self.text = ""      # extract text
        self.meta = ""      # extract metadata

        self.isRestorable = self.CONST_DOCUMENT_UNRESTORABLE
        self.isEncrypted = self.CONST_DOCUMENT_NO_ENCRYPTED

    def __enter__(self):
        raise NotImplementedError

    def __exit__(self):
        raise NotImplementedError


    def __parse_xls_normal__(self):

        RECORD_HEADER_SIZE = 4
        records = []
        # 원하는 스트림 f에 모두 읽어오기
        temp = self.fp.open('Workbook').read()
        f = bytearray(temp)
        # 스트림 내부 모두 파싱해서 데이터 출력
        tempOffset = 0

        while tempOffset < len(f):
            dic = {}
            dic['offset'] = tempOffset
            dic['type'] = struct.unpack('<h', f[tempOffset: tempOffset + 0x02])[0]
            dic['length'] = struct.unpack('<h', f[tempOffset + 0x02: tempOffset + 0x04])[0]
            dic['data'] = f[tempOffset + RECORD_HEADER_SIZE: tempOffset + RECORD_HEADER_SIZE + dic['length']]
            tempOffset = tempOffset + RECORD_HEADER_SIZE + dic['length']
            records.append(dic)


        # Continue marker
        for record in records:
            if record['type'] == 0xFC:
                sstNum = records.index(record)
                sstOffset = record['offset']
                sstLen = record['length']
            if record['type'] == 0x3C:
                f[record['offset']:record['offset']+4] = bytearray(b'\xAA\xAA\xAA\xAA')


        cntStream = sstOffset + 4
        cstTotal = struct.unpack('<i', f[cntStream : cntStream + 4])[0]
        cstUnique = struct.unpack('<i', f[cntStream + 4: cntStream + 8])[0]
        cntStream += 8


        for i in range(0, cstUnique):
            string = ""
            if(cntStream > len(f)):
                break
            # if start is Continue
            if f[cntStream: cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                cntStream += 4

            cch = struct.unpack('<H', f[cntStream: cntStream + 2])[0]  ### 문자열 길이
            cntStream += 2
            flags = f[cntStream]  ### 플래그를 이용해서 추가적 정보 확인
            cntStream += 1

            if cch == 0x00 and flags == 0x00:
                continue

            if cch == 0x00:
                break

            if flags & 0x02 or flags >= 0x10:
                break


            if (flags & 0b00000001 == 0b00000001):
                fHighByte = 0x01
            else:
                fHighByte = 0x00

            if (flags & 0b00000100 == 0b00000100):
                fExtSt = 0x01
            else:
                fExtSt = 0x00

            if (flags & 0b00001000 == 0b00001000):
                fRichSt = 0x01
            else:
                fRichSt = 0x00

            if fRichSt == 0x01:
                cRun = struct.unpack('<H', f[cntStream: cntStream + 2])[0]
                cntStream += 2

            if fExtSt == 0x01:
                cbExtRst = struct.unpack('<I', f[cntStream: cntStream + 4])[0]
                cntStream += 4

            if fHighByte == 0x00:  ### Ascii
                bAscii = True
                for j in range(0, cch):
                    if f[cntStream : cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                        if f[cntStream + 4] == 0x00 or f[cntStream + 4] == 0x01:
                            cntStream += 4

                            if f[cntStream] == 0x00:
                                bAscii = True
                            elif f[cntStream] == 0x01:
                                bAscii = False

                            cntStream += 1

                    if bAscii == True:
                        try:
                            string += str(bytes([f[cntStream]]).decode("ascii"))
                            cntStream += 1
                        except UnicodeDecodeError:
                            cntStream += 1
                            continue

                    elif bAscii == False:
                        try:
                            string += str(f[cntStream: cntStream + 2].decode("utf-16"))
                            cntStream += 2
                        except UnicodeDecodeError:
                            cntStream += 2
                            continue

            elif fHighByte == 0x01:  ### Unicode
                bAscii = False
                for j in range(0, cch):

                    if f[cntStream : cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                        if f[cntStream + 4] == 0x00 or f[cntStream + 4] == 0x01:
                            cntStream += 4

                            if f[cntStream] == 0x00:
                                bAscii = True
                            elif f[cntStream] == 0x01:
                                bAscii = False

                            cntStream += 1


                    if bAscii == True:
                        try :
                            string += str(bytes([f[cntStream]]).decode("ascii"))
                            cntStream += 1
                        except UnicodeDecodeError:
                            cntStream += 1
                            continue

                    elif bAscii == False:
                        try :
                            string += str(f[cntStream: cntStream + 2].decode("utf-16"))
                            cntStream += 2
                        except UnicodeDecodeError:
                            cntStream += 2
                            continue

            print(str(i) + " : " + string)

            if fRichSt == 0x01:
                if f[cntStream: cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                    cntStream += 4
                cntStream += int(cRun) * 4

            if fExtSt == 0x01:
                for i in range(0, cbExtRst):
                    if cntStream > len(f):
                        break

                    if f[cntStream: cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                        if i + 4 <= cbExtRst:
                            cntStream += 4

                    cntStream += 1

    def __parse_xls_damaged__(self):
        test = bytearray(self.fp.read())
        tempOffset = 0
        globalStreamOffset = 0
        while tempOffset < len(test):
            if test[tempOffset:tempOffset+8] == b'\x09\x08\x10\x00\x00\x06\x05\x00':
                globalStreamOffset = tempOffset
                break
            tempOffset += 0x80

        print(globalStreamOffset)
        f = test[globalStreamOffset:]

        RECORD_HEADER_SIZE = 4
        records = []
        # 스트림 내부 모두 파싱해서 데이터 출력
        tempOffset = 0

        while tempOffset < len(f):
            dic = {}
            dic['offset'] = tempOffset
            dic['type'] = struct.unpack('<h', f[tempOffset: tempOffset + 0x02])[0]
            if dic['type'] >= 4200 or dic['type'] <= 6:
                break
            dic['length'] = struct.unpack('<h', f[tempOffset + 0x02: tempOffset + 0x04])[0]
            if dic['length'] >= 8225:
                break
            dic['data'] = f[tempOffset + RECORD_HEADER_SIZE: tempOffset + RECORD_HEADER_SIZE + dic['length']]
            tempOffset = tempOffset + RECORD_HEADER_SIZE + dic['length']
            records.append(dic)

        bSST = False
        # Continue marker
        for record in records:
            if record['type'] == 0xFC:
                sstOffset = record['offset']
                bSST = True
            if record['type'] == 0x3C:
                f[record['offset']:record['offset']+4] = bytearray(b'\xAA\xAA\xAA\xAA')

        if bSST == False:
            return self.CONST_ERROR


        cntStream = sstOffset + 4
        cstTotal = struct.unpack('<i', f[cntStream : cntStream + 4])[0]
        cstUnique = struct.unpack('<i', f[cntStream + 4: cntStream + 8])[0]
        cntStream += 8


        for i in range(0, cstUnique):
            string = ""
            if(cntStream > len(f)):
                break
            # if start is Continue
            if f[cntStream: cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                cntStream += 4

            cch = struct.unpack('<H', f[cntStream: cntStream + 2])[0]  ### 문자열 길이
            cntStream += 2
            flags = f[cntStream]  ### 플래그를 이용해서 추가적 정보 확인
            cntStream += 1

            if cch == 0x00 and flags == 0x00:
                continue

            if cch == 0x00:
                break

            if flags & 0x02 or flags >= 0x10:
                break


            if (flags & 0b00000001 == 0b00000001):
                fHighByte = 0x01
            else:
                fHighByte = 0x00

            if (flags & 0b00000100 == 0b00000100):
                fExtSt = 0x01
            else:
                fExtSt = 0x00

            if (flags & 0b00001000 == 0b00001000):
                fRichSt = 0x01
            else:
                fRichSt = 0x00

            if fRichSt == 0x01:
                cRun = struct.unpack('<H', f[cntStream: cntStream + 2])[0]
                cntStream += 2

            if fExtSt == 0x01:
                cbExtRst = struct.unpack('<I', f[cntStream: cntStream + 4])[0]
                cntStream += 4

            if fHighByte == 0x00:  ### Ascii
                bAscii = True
                for j in range(0, cch):
                    if f[cntStream : cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                        if f[cntStream + 4] == 0x00 or f[cntStream + 4] == 0x01:
                            cntStream += 4

                            if f[cntStream] == 0x00:
                                bAscii = True
                            elif f[cntStream] == 0x01:
                                bAscii = False

                            cntStream += 1

                    if bAscii == True:
                        try:
                            string += str(bytes([f[cntStream]]).decode("ascii"))
                            cntStream += 1
                        except UnicodeDecodeError:
                            cntStream += 1
                            continue

                    elif bAscii == False:
                        try:
                            string += str(f[cntStream: cntStream + 2].decode("utf-16"))
                            cntStream += 2
                        except UnicodeDecodeError:
                            cntStream += 2
                            continue

            elif fHighByte == 0x01:  ### Unicode
                bAscii = False
                for j in range(0, cch):

                    if f[cntStream : cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                        if f[cntStream + 4] == 0x00 or f[cntStream + 4] == 0x01:
                            cntStream += 4

                            if f[cntStream] == 0x00:
                                bAscii = True
                            elif f[cntStream] == 0x01:
                                bAscii = False

                            cntStream += 1


                    if bAscii == True:
                        try :
                            string += str(bytes([f[cntStream]]).decode("ascii"))
                            cntStream += 1
                        except UnicodeDecodeError:
                            cntStream += 1
                            continue

                    elif bAscii == False:
                        try :
                            string += str(f[cntStream: cntStream + 2].decode("utf-16"))
                            cntStream += 2
                        except UnicodeDecodeError:
                            cntStream += 2
                            continue

            print(str(i) + " : " + string)

            if fRichSt == 0x01:
                if f[cntStream: cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                    cntStream += 4
                cntStream += int(cRun) * 4

            if fExtSt == 0x01:
                for i in range(0, cbExtRst):
                    if cntStream > len(f):
                        break

                    if f[cntStream: cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                        if i + 4 <= cbExtRst:
                            cntStream += 4

                    cntStream += 1


    def __doc_extra_filter__(self, string, uFilteredTextLen):
        i = 0
        j = 0
        k = 0

        # 1.        첫        부분의        공백        문자        모두        제거
        # 2.        공백        문자가        2        개        이상인        경우에        1        개로        만들자
        # 3.        개행        문자가        2        개        이상인        경우에        1        개로        만들자
        # 4.        Filtering

        uBlank = 0x0020   # ASCII Blank
        uBlank2 = 0x00A0   # Unicode Blank
        uNewline = 0x000A   # Line Feed
        uNewline2 = 0x000D
        uNewline3 = 0x0004
        uNewline4 = 0x0003
        uSection = 0x0001
        uSection2 = 0x0002
        uSection3 = 0x0005
        uSection4 = 0x0007
        uSection5 = 0x0008
        uSection6 = 0x0015
        uSection7 = 0x000C
        uSection8 = 0x000B
        uSection9 = 0x0014
        uTrash = 0x0000
        uCaption = [0x0053, 0x0045, 0x0051]
        uCaption2 = [0x0041, 0x0052, 0x0041, 0x0042, 0x0049, 0x0043, 0x0020, 0x0014]
        uHyperlink = [0x0048, 0x0059, 0x0050, 0x0045, 0x0052, 0x004C, 0x0049, 0x004E, 0x004B]
        uToc = [0x0054, 0x004F]
        uPageref = [0x0050, 0x0041, 0x0047, 0x0045, 0x0052, 0x0045, 0x0046]
        uIndex = [0x0049, 0x004E, 0x0044, 0x0045, 0x0058]
        uEnd = [0x0020, 0x0001, 0x0014]
        uEnd2 = [0x0020, 0x0014]
        uEnd3 = [0x0020, 0x0015]
        uEnd4 = 0x0014
        uEnd5 = [0x0001, 0x0014]
        uEnd6 = 0x0015
        uHeader = 0x0013
        uChart = [0x0045, 0x004D, 0x0042, 0x0045, 0x0044]
        uShape = [0x0053, 0x0048, 0x0041, 0x0050, 0x0045]
        uPage = [0x0050, 0x0041, 0x0047, 0x0045]
        uDoc = [0x0044, 0x004F, 0x0043]
        uStyleref = [0x0053, 0x0054, 0x0059, 0x004C, 0x0045, 0x0052, 0x0045, 0x0046]
        uTitle_text = [0x0054, 0x0049, 0x0054, 0x004C, 0x0045]
        uDate = [0x0049, 0x0046, 0x0020, 0x0044, 0x0041, 0x0054, 0x0045]
        FilteredText = string

        for i in range(0, uFilteredTextLen, 2):
            if i == 0:
                k = 0
                temp = struct.unpack('<H', FilteredText[0:2])[0]
                while (temp == uBlank or temp == uBlank2 or temp == uNewline or temp == uNewline2 or
                       temp == uNewline3 or temp == uNewline4) :

                    FilteredText = FilteredText[:k] + FilteredText[k+2:]
                    uFilteredTextLen -= 2

                    if (len(FilteredText) == 0):
                        break

            if len(FilteredText) == 0:
                break

        return 0

    def __parse_doc_normal__(self):
        word_document = bytearray(self.fp.open('WordDocument').read())  # byteWD
        one_table = b''
        zero_table = b''
        try :
            one_table = bytearray(self.fp.open('1Table').read())
        except compoundfiles.errors.CompoundFileNotFoundError:
            print("1Table is not exist.")

        try :
            zero_table = bytearray(self.fp.open('0Table').read())
        except compoundfiles.errors.CompoundFileNotFoundError:
            print("0Table is not exist.")


        if len(one_table) == 0 and len(zero_table) == 0:
            return Compound.CONST_ERROR

        # Extract doc Text
        ccpText = b''
        fcClx = b''
        lcbClx = b''
        aCP = b''
        aPcd = b''
        fcCompressed = b''
        Clx = b''
        byteTable = b''
        ccpTextSize = 0
        fcClxSize = 0
        lcbClxSize = 0
        ClxSize = 0
        string = b''
        CONST_FCFLAG = 1073741824		# 0x40000000
        CONST_FCINDEXFLAG = 1073741823	# 0x3FFFFFFF
        i = 0
        j = 0

        # Check Encrypted
        uc_temp = word_document[11]
        uc_temp = uc_temp & 1

        if uc_temp == 1:
            return Compound.CONST_ERROR


        # 0Table 1Table
        is0Table = word_document[11] & 2

        if is0Table == 0 :
            byteTable = zero_table
        else:
            byteTable = one_table


        # Get cppText in FibRgLw
        ccpText = word_document[0x4C:0x50]
        ccpTextSize = struct.unpack('<I', ccpText)[0]

        if (ccpTextSize == 0):
            return Compound.CONST_ERROR


        # Get fcClx in FibRgFcLcbBlob
        fcClx = word_document[0x1A2:0x1A6]
        fcClxSize = struct.unpack('<I', fcClx)[0]

        if (fcClxSize == 0):
            return Compound.CONST_ERROR


        # Get lcbClx in FibRgFcLcbBlob
        lcbClx = word_document[0x1A6:0x1AA]  
        lcbClxSize = struct.unpack('<I', lcbClx)[0]

        if (lcbClxSize == 0):
            return Compound.CONST_ERROR


        # Get Clx
        Clx = byteTable[fcClxSize : fcClxSize + lcbClxSize]

        if Clx[0] == 0x01:
            cbGrpprl = Clx[1:3]
            Clx = byteTable[fcClxSize + cbGrpprl + 3 : (fcClxSize + cbGrpprl + 3) + lcbClxSize - cbGrpprl + 3]
        if Clx[0] != 0x02:
            return Compound.CONST_ERROR

        ClxSize = struct.unpack('<I', Clx[1:5])[0]




        ClxIndex = 5
        PcdCount = 0
        aCPSize = []
        fcFlag = 0
        fcIndex = 0
        fcSize = 0
        encodingFlag = False
        
        PcdCount = int(((ClxSize / 4) / 3)) + 1
        
        for k in range(0, PcdCount):
            aCp = Clx[ClxIndex:ClxIndex+4]
            aCPSize.append(struct.unpack('<I', aCp[0:4])[0])
            ClxIndex += 4
            
        PcdCount -= 1

        ### Filtering

         
        uBlank = 0x0020   # ASCII Blank
        uBlank2 = 0x00A0   # Unicode Blank
        uNewline = 0x000A   # Line Feed
        uNewline2 = 0x000D
        uNewline3 = 0x0004
        uNewline4 = 0x0003
        uSection = 0x0001
        uSection2 = 0x0002
        uSection3 = 0x0005
        uSection4 = 0x0007
        uSection5 = 0x0008
        uSection6 = 0x0015
        uSection7 = 0x000C
        uSection8 = 0x000B
        uSection9 = 0x0014
        uTrash = 0x0000
        uCaption = [0x0053, 0x0045, 0x0051]
        uCaption2 = [0x0041, 0x0052, 0x0041, 0x0042, 0x0049, 0x0043, 0x0020, 0x0014]
        uHyperlink = [0x0048, 0x0059, 0x0050, 0x0045, 0x0052, 0x004C, 0x0049, 0x004E, 0x004B]
        uToc = [0x0054, 0x004F]
        uPageref = [0x0050, 0x0041, 0x0047, 0x0045, 0x0052, 0x0045, 0x0046]
        uIndex = [0x0049, 0x004E, 0x0044, 0x0045, 0x0058]
        uEnd = [0x0020, 0x0001, 0x0014]
        uEnd2 = [0x0020, 0x0014]
        uEnd3 = [0x0020, 0x0015]
        uEnd4 = 0x0014
        uEnd5 = [0x0001, 0x0014]
        uEnd6 = 0x0015
        uHeader = 0x0013
        uChart = [0x0045, 0x004D, 0x0042, 0x0045, 0x0044]
        uShape = [0x0053, 0x0048, 0x0041, 0x0050, 0x0045]
        uPage = [0x0050, 0x0041, 0x0047, 0x0045]
        uDoc = [0x0044, 0x004F, 0x0043]
        uStyleref = [0x0053, 0x0054, 0x0059, 0x004C, 0x0045, 0x0052, 0x0045, 0x0046]
        uTitle_text = [0x0054, 0x0049, 0x0054, 0x004C, 0x0045]
        uDate = [0x0049, 0x0046, 0x0020, 0x0044, 0x0041, 0x0054, 0x0045]

        ### Filtering targets: 0x0001 ~ 0x0017(0x000A Line Feed skipped)
        uTab = 0x0009 # Horizontal Tab
        uSpecial = 0xF0
        bFullScanA = False
        bFullScanU = False # if the size info is invalid, then the entire range will be scanned.
        tempPlus = 0

        for i in range(0, PcdCount):
            aPcd = Clx[ClxIndex:ClxIndex + 8]
            fcCompressed = aPcd[2:6]

            fcFlag = struct.unpack('<I', fcCompressed[0:4])[0]

            if CONST_FCFLAG == (fcFlag & CONST_FCFLAG):
                encodingFlag = True                 # 8-bit ANSI
            else:
                encodingFlag = False                # 16-bit Unicode

            fcIndex = fcFlag & CONST_FCINDEXFLAG






            if encodingFlag == True:                # 8-bit ANSI
                fcIndex = int(fcIndex / 2)
                fcSize = aCPSize[i+1] - aCPSize[i]

                if len(word_document) < fcIndex + fcSize + 1:
                    if bFullScanA == False and len(word_document) > fcIndex:
                        fcSize = len(word_document) - fcIndex - 1
                        bFullScanA = True
                    else:
                        ClxIndex += 8
                        continue


                ASCIIText = word_document[fcIndex:fcIndex + fcSize]
                UNICODEText = ASCIIText.decode('utf-8')

                for i in range(0, len(UNICODEText), 2):
                    temp = struct.unpack('<H', UNICODEText[i:i + 2])[0]
                    if ( temp == uSection2 or temp == uSection3[0:2] or temp == uSection4 or
                        temp == uSection5 or temp == uSection7 or temp == uSection8 or
                        UNICODEText[i + 1] == uSpecial or temp == uTrash[0:2] ) :
                        continue

                    if ( temp == uNewline or temp == uNewline2 or temp == uNewline3 or temp == uNewline4 ):
                        string += bytes([UNICODEText[i]])
                        string += bytes([UNICODEText[i + 1]])

                        for j in range(i + 2, len(UNICODEText) * 2, 2):
                            temp = struct.unpack('H', UNICODEText[j:j + 2])[0]
                            if ( temp == uSection2 or temp == uSection3 or temp == uSection4 or
                                 temp == uSection5 or temp == uSection7 or temp == uSection8 or
                                 temp == uBlank or temp == uBlank2 or temp == uNewline or
                                 temp == uNewline2 or temp == uNewline3 or temp == uNewline4 or
                                 temp == uTab or UNICODEText[j + 1] == uSpecial ):
                                continue
                            else:
                                i = j
                                break
                        if j >= len(UNICODEText) * 2 :
                            break
                    elif ( temp == uBlank or temp == uBlank2 or temp == uTab ):

                        string += bytes([UNICODEText[i]])
                        string += bytes([UNICODEText[i + 1]])

                        for j in range(i+2, len(UNICODEText) * 2, 2):
                            if (temp == uSection2 or temp == uSection3 or temp == uSection4 or
                                    temp == uSection5 or temp == uSection7 or temp == uSection8 or
                                    temp == uBlank or temp == uBlank2 or temp == uTab or UNICODEText[j + 1] == uSpecial):
                                continue
                            else:

                                i = j
                                break


                        if (j >= len(UNICODEText) * 2):
                            break

                    string += bytes([UNICODEText[i]])
                    string += bytes([UNICODEText[i + 1]])


            elif encodingFlag == False :          ### 16-bit Unicode
                fcSize = 2 * (aCPSize[i + 1] - aCPSize[i])

                if(len(word_document) < fcIndex + fcSize + 1):   # Invalid structure - size info is invalid (large) => scan from fcIndex to last
                    if (bFullScanU == False and len(word_document) > fcIndex):
                        fcSize = len(word_document) - fcIndex -1
                        bFullScanU = True
                    else:
                        ClxIndex = ClxIndex + 8
                        continue

                while i < fcSize:
                    temp = struct.unpack('<H', word_document[fcIndex + i : fcIndex + i + 2])[0]
                    if ( temp == uSection2 or temp == uSection3 or
                        temp == uSection4 or temp == uSection5 or
                        temp == uSection7 or temp == uSection8 or
                         word_document[fcIndex + i + 1] == uSpecial or temp == uTrash ):
                        continue

                    if ( temp == uNewline or temp == uNewline2 or temp == uNewline3 or temp == uNewline4 ):

                        if ( word_document[fcIndex + i] == 0x0d ):
                            string += b'\x0a'
                            string += bytes([word_document[fcIndex + i + 1]])
                        else :
                            string += bytes([word_document[fcIndex + i]])
                            string += bytes([word_document[fcIndex + i + 1]])

                        for j in range(i+2, fcSize, 2):
                            temp2 = struct.unpack('<H', word_document[fcIndex + j : fcIndex + j + 2])[0]
                            if ( temp2 == uSection2 or temp2 == uSection3 or temp2 == uSection4 or
                                temp2 == uSection5 or temp2 == uSection7 or temp2 == uSection8 or
                                temp2 == uBlank or temp2 == uBlank2 or temp2 == uNewline or temp2 == uNewline2 or
                                temp2 == uNewline3 or temp2 == uNewline4 or temp2 == uTab or word_document[fcIndex + j + 1] == uSpecial ) :
                                continue
                            else :
                                i = j
                                break

                        if j >= fcSize:
                            break

                    elif temp == uBlank or temp == uBlank2 or temp == uTab :
                        string += bytes([word_document[fcIndex + i]])
                        string += bytes([word_document[fcIndex + i + 1]])

                        for j in range(i+2, fcSize, 2):
                            temp2 = struct.unpack('<H', word_document[fcIndex + j : fcIndex + j + 2])[0]
                            if ( temp2 == uSection2 or temp2 == uSection3 or temp2 == uSection4 or
                                temp2 == uSection5 or temp2 == uSection7 or temp2 == uSection8 or
                                temp2 == uBlank or temp2 == uBlank2 or temp2 == uTab or word_document[fcIndex + j + 1] == uSpecial ) :
                                continue
                            else :
                                i = j
                                break

                        if j >= fcSize:
                            break


                    string += bytes([word_document[fcIndex + i]])
                    string += bytes([word_document[fcIndex + i + 1]])
                    i += 2
                    print("while", i)

            ClxIndex += 8

        # Test
        print("test end")

        uFilteredTextLen = self.__doc_extra_filter__(string, len(string))









    def __parse_doc_damaged__(self):
        raise NotImplementedError

    def __parse_ppt_normal__(self):
        raise NotImplementedError

    def __parse_ppt_damaged__(self):
        raise NotImplementedError


    def parse(self):

        if self.fileType == "xls" :
            result = self.parse_xls()
        elif self.fileType == "ppt" :
            result = self.parse_ppt()
        elif self.fileType == "doc" :
            result = self.parse_doc()

        """
        if result == self.CONST_SUCCESS:
            return self.CONST_SUCCESS
        elif result == self.CONST_ERROR:
            return self.CONST_ERROR
        """
        #self.parse_xls()
        #self.parse_summaryinfo()

    def parse_xls(self):

        if self.isDamaged == self.CONST_DOCUMENT_NORMAL:
            self.__parse_xls_normal__()
        elif self.isDamaged == self.CONST_DOCUMENT_DAMAGED:
            self.__parse_xls_damaged__()



    def parse_ppt(self):
        raise NotImplementedError

    def parse_doc(self):
        if self.isDamaged == self.CONST_DOCUMENT_NORMAL:
            self.__parse_doc_normal__()
        elif self.isDamaged == self.CONST_DOCUMENT_DAMAGED:
            self.__parse_doc_damaged__()


    def parse_summaryinfo(self):
        records = []
        # Open SummaryInformation Stream
        f = self.fp.open('\x05SummaryInformation').read()

        startOffset = struct.unpack('<i', f[0x2C: 0x30])[0]
        tempOffset = startOffset

        # Store Records
        length = struct.unpack('<i', f[tempOffset: tempOffset + 0x04])[0]
        recordCount = struct.unpack('<i', f[tempOffset + 0x04: tempOffset + 0x08])[0]
        tempOffset += 0x08
        for i in range(0, recordCount):
            dict = {}
            dict['type'] = struct.unpack('<i', f[tempOffset: tempOffset + 0x04])[0]
            dict['offset'] = struct.unpack('<i', f[tempOffset + 0x04: tempOffset + 0x08])[0]
            records.append(dict)
            tempOffset += 0x08

        # Print Records
        for record in records:

            # Title
            if record['type'] == 0x02:
                entryLength = \
                struct.unpack('<i', f[record['offset'] + startOffset + 4: record['offset'] + startOffset + 8])[0]
                entryData = f[record['offset'] + startOffset + 8: record['offset'] + startOffset + 8 + entryLength]
                print(entryData.decode('euc-kr'))


            # Subject
            elif record['type'] == 0x03:
                entryLength = \
                struct.unpack('<i', f[record['offset'] + startOffset + 4: record['offset'] + startOffset + 8])[0]
                entryData = f[record['offset'] + startOffset + 8: record['offset'] + startOffset + 8 + entryLength]
                print(entryData.decode('euc-kr'))

            # Author
            elif record['type'] == 0x04:
                entryLength = \
                struct.unpack('<i', f[record['offset'] + startOffset + 4: record['offset'] + startOffset + 8])[0]
                entryData = f[record['offset'] + startOffset + 8: record['offset'] + startOffset + 8 + entryLength]
                print(entryData.decode('euc-kr'))

            # LastAuthor
            elif record['type'] == 0x08:
                entryLength = \
                struct.unpack('<i', f[record['offset'] + startOffset + 4: record['offset'] + startOffset + 8])[0]
                entryData = f[record['offset'] + startOffset + 8: record['offset'] + startOffset + 8 + entryLength]
                print(entryData.decode('euc-kr'))

            # AppName
            elif record['type'] == 0x12:
                entryLength = \
                struct.unpack('<i', f[record['offset'] + startOffset + 4: record['offset'] + startOffset + 8])[0]
                entryData = f[record['offset'] + startOffset + 8: record['offset'] + startOffset + 8 + entryLength]
                print(entryData.decode('euc-kr'))

            # LastPrintedtime
            elif record['type'] == 0x0B:
                entryTimeData = struct.unpack('<q', f[record['offset'] + startOffset + 4: record['offset'] + startOffset + 12])[0] / 1e8
                print(datetime.datetime.fromtimestamp(entryTimeData).strftime('%Y-%m-%d %H:%M:%S.%f'))

            # Createtime
            elif record['type'] == 0x0C:
                entryTimeData = struct.unpack('<q', f[record['offset'] + startOffset + 4: record['offset'] + startOffset + 12])[0] / 1e8
                print(datetime.datetime.fromtimestamp(entryTimeData).strftime('%Y-%m-%d %H:%M:%S.%f'))

            # LastSavetime
            elif record['type'] == 0x0D:
                entryTimeData = struct.unpack('<q', f[record['offset'] + startOffset + 4: record['offset'] + startOffset + 12])[0] / 1e8
                print(datetime.datetime.fromtimestamp(entryTimeData).strftime('%Y-%m-%d %H:%M:%S.%f'))
