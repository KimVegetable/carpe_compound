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
            self.fp = compoundfiles.CompoundFileReader(filePath)
            print("File exist!!")
        else:
            self.fp = None
            print("File doesn't exist.")

        self.fileSize = os.path.getsize(filePath)
        self.fileName = os.path.basename(filePath)
        self.filePath = filePath
        self.fileType = os.path.splitext(filePath)[1][1:]   # delete '.' in '.xls' r
        self.text = ""      # extract text
        self.meta = ""      # extract metadata

        self.isDamaged = self.CONST_DOCUMENT_NORMAL
        self.isRestorable = self.CONST_DOCUMENT_UNRESTORABLE
        self.isEncrypted = self.CONST_DOCUMENT_NO_ENCRYPTED

    def __enter__(self):
        raise NotImplementedError

    def __exit__(self):
        raise NotImplementedError

    def parse(self):
        """
        if self.fileType == "xls" :
            result = self.parse_xls()
        elif self.fileType == "ppt" :
            result = self.parse_ppt()
        elif self.fileType == "doc" :
            result = self.parse_doc()


        if result == self.CONST_SUCCESS:
            return self.CONST_SUCCESS
        elif result == self.CONST_ERROR:
            return self.CONST_ERROR
        """
        self.parse_xls()
        self.parse_summaryinfo()


    def parse_xls(self):
        RECORD_HEADER_SIZE = 4

        records = []

        # 원하는 스트림 f에 모두 읽어오기
        f = self.fp.open('Workbook').read()

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
                f[record['offset']:record['offset']+4] = b'\xAA\xAA\xAA\xAA'


        cntStream = sstOffset
        cstTotal = struct.unpack('<i', f[cntStream : cntStream + 4])[0]
        cstUnique = struct.unpack('<i', f[cntStream + 4: cntStream + 8])[0]
        cntStream += 8

        string = b''
        for i in range(0, cstUnique):

            cch = struct.unpack('<h', f[cntStream: cntStream + 2])[0]  ### 문자열 길이
            cntStream += 2
            flags = f[cntStream]  ### 플래그를 이용해서 추가적 정보 확인
            cntStream += 1

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
                cRun = struct.unpack('<h', f[cntStream: cntStream + 2])[0]
                cntStream += 2
            if fExtSt == 0x01:
                cntStream += 4

            if fHighByte == 0x00:  ### Ascii
                for j in range(0, cch):
                    bAscii = True

                    if f[cntStream : cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                        if f[cntStream + 4] == 0x00 or f[cntStream + 4] == 0x01:
                            cntStream += 4

                            if f[cntStream] == 0x00:
                                bAscii = True
                            elif f[cntStream] == 0x01:
                                bAscii = False

                            cntStream += 1

                    if bAscii == True:
                        string += f[cntStream]
                        cntStream += 1

                    elif bAscii == False:
                        string += f[cntStream : cntStream + 2]
                        cntStream += 2

            elif fHighByte == 0x01:  ### Unicode
                for j in range(0, cch):
                    bAscii = False

                    if f[cntStream: cntStream + 4] == b'\xAA\xAA\xAA\xAA':
                        if f[cntStream + 4] == 0x00 or f[cntStream + 4] == 0x01:
                            cntStream += 4

                            if f[cntStream] == 0x00:
                                bAscii = True
                            elif f[cntStream] == 0x01:
                                bAscii = False

                            cntStream += 1

                    if bAscii == True:
                        string += f[cntStream]
                        cntStream += 1

                    elif bAscii == False:
                        string += f[cntStream: cntStream + 2]
                        cntStream += 2
            string += b'\n'
        print(string)








        """
            if fHighByte == 0x01:  ### 유니코드
                string += f[cntStream: cntStream + cch * 2]
                cntStream += cch * 2
            else:  ### 아스키
                string += f[cntStream: cntStream + cch]
                cntStream += cch




            if fRichSt == 0x01:
                cntStream += int(cRun) * 4
            if fExtSt == 0x01:
                cntStream += 16



            if fHighByte == 0x01:
                print(str(i) + " " + str(string, "utf-16"))
            else:
                print(str(i) + " " + str(string))
            string += b'\n'
            # print(dict['string'].decode("utf-8"))
        """


#        for i in range(records[sstNum]['offset'], records[sstNum]['offset'] + records[sstNum]['length']):
            #print(hex(f[i]))


        """
        # 파일에 입력
        for record in records:
            # print(record['data'].hex())

            if record['type'] == 0xFC:
                tempOffset = 0
                SST_records = []
                SST = {}

                SST['cstTotal'] = struct.unpack('<i', SSTData[tempOffset: tempOffset + 4])[0]
                SST['cstUnique'] = struct.unpack('<i', SSTData[tempOffset + 4: tempOffset + 8])[0]
                tempOffset += 8

                for i in range(0, SST['cstUnique']):

                    dict = {}
                    dict['cch'] = struct.unpack('<h', SSTData[tempOffset: tempOffset + 2])[0]  ### 문자열 길이
                    tempOffset += 2
                    dict['flags'] = SSTData[tempOffset]  ### 플래그를 이용해서 추가적 정보 확인
                    tempOffset += 1

                    if (dict['flags'] & 0b00000001 == 0b00000001):
                        dict['fHighByte'] = 0x01
                    else:
                        dict['fHighByte'] = 0x00

                    if (dict['flags'] & 0b00000100 == 0b00000100):
                        dict['fExtSt'] = 0x01
                    else:
                        dict['fExtSt'] = 0x00

                    if (dict['flags'] & 0b00001000 == 0b00001000):
                        dict['fRichSt'] = 0x01
                    else:
                        dict['fRichSt'] = 0x00

                    if dict['fRichSt'] == 0x01:
                        dict['cRun'] = struct.unpack('<h', SSTData[tempOffset: tempOffset + 2])[0]
                        tempOffset += 2

                    if dict['fExtSt'] == 0x01:
                        tempOffset += 4

                    if dict['fHighByte'] == 0x01:  ### 유니코드
                        dict['string'] = SSTData[tempOffset: tempOffset + dict['cch'] * 2]
                        tempOffset += dict['cch'] * 2
                    else:  ### 아스키
                        dict['string'] = SSTData[tempOffset: tempOffset + dict['cch']]
                        tempOffset += dict['cch']

                    if dict['fRichSt'] == 0x01:
                        tempOffset += int(dict['cRun']) * 4
                    if dict['fExtSt'] == 0x01:
                        tempOffset += 16


                    if dict['fHighByte'] == 0x01:
                        print(str(i) + " " + str(dict['string'], "utf-16"))
                    else:
                        print(str(i) + " " + str(dict['string']))

                    # print(dict['string'].decode("utf-8"))
                    SST_records.append(dict)
        """

    def parse_ppt(self):
        raise NotImplementedError

    def parse_doc(self):
        raise NotImplementedError

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
