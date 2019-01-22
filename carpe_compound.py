# carpe_compound.py

import compoundfiles
import os


class COMPOUND:

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

    def parse_xls(self):
        raise NotImplementedError

    def parse_ppt(self):
        raise NotImplementedError

    def parse_doc(self):
        raise NotImplementedError
