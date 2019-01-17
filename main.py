from carpe_compound import COMPOUND
import sys
import os
import compoundfiles
from mariadb import MariaDB


def main(filePath):
    """
    with MariaDB(user='root', password='toor') as maria:
        maria.query("CREATE DATABASE carpe_vegetable")
    """

    object = COMPOUND(filePath)

    if object.fp == None:
        return

    ### Test Code ###
    print("fileSize : " + str(object.fileSize))
    print("fileName : " + str(object.fileName))
    print("fileType : " + str(object.fileType))
    print("filePath : " + str(object.filePath))

    object.ParseDocument()




    object.fp.close()

if __name__ == "__main__":
    main(sys.argv[1])