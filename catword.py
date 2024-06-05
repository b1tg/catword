# -*- coding: utf-8 -*-
import sys, os, struct, re
import olefile, hexdump
from zipfile import is_zipfile, ZipFile

def parse_docx(file_name):
    zf = ZipFile(file_name)
    for sub_file in zf.namelist():
        if sub_file == 'word/document.xml':
            with zf.open(sub_file) as file_handle:
                content = file_handle.read()
                content = content.decode("utf8")
                print(content)
                print("*"*20)
                if "</w:t>" in content:
                    # remove xml tags to fix fonts error
                    # content = content.replace('w:lineRule="auto"/>', \
                    #     "/><w:t>CRCR</w:t>")
                    # content = content.replace("<w:p", "<w:x")
                    content = content.replace("</w:p>", "</w:p><w:t>CRCR</w:t>")
                    content = re.sub("</?w:[^>]*>", "", content) 
                    # content = content.split("\r\n")
                    content = content.replace("CRCR", "\n")
                    return content
    return ""
def parse_doc(file_name):
    ole = olefile.OleFileIO(file_name)
    for path_parts in ole.listdir():
        # print(path_parts)
        if path_parts[-1] in("WordDocument",  ):
            stream_path = "/".join(path_parts)
            stream = ole.openstream(stream_path)
            W = stream.read()
        elif path_parts[-1] == "1Table":
            stream_path = "/".join(path_parts)
            stream = ole.openstream(stream_path)
            T = stream.read()

    textinfo_off = struct.unpack("<H", W[0x01a2:0x01a4])[0]
    # print(hex(textinfo_off))

    # hexdump.hexdump(T[textinfo_off:textinfo_off+0x30])
    # T[0x2997:] = 
    # findtext.c[233]:
    # 000: 02 1c 00 00 00 00 00 00 00 00 01 00 00 b5 01 00 
    # 010: 00 10 01 00 08 00 00 00 00 10 01 00 0e 00 00 00 
    # 020: 00 

    ulLen = struct.unpack("<I", T[textinfo_off+1:textinfo_off+5])[0]
    lPieces = int((ulLen-4)/12)
    # print("ulLen: 0x%x" %(ulLen))
    # print("lPieces: 0x%x" %(lPieces))

    off = textinfo_off+5
    res = ""
    for lIndex in range(0, lPieces):
        i = off+(lPieces+1)*4 + lIndex*8 +2
        text_off = struct.unpack("<I", T[i:i+4])[0]
        i = off+(lIndex+1)*4
        j = off+lIndex*4
        text_len = struct.unpack("<I", T[i:i+4])[0] - struct.unpack("<I", T[j:j+4])[0]
        # text_len = text_len*2
        # print("text_off: 0x%x, text_len: 0x%x" %(text_off, text_len))
        data = W[text_off:text_len*2+text_off]
        # hexdump.hexdump(data)

        # for i in range(0, len(data), 2):
        #     print("data:%s" %(data[i:i+2].decode("utf16")))
        #     print(data[i:i+2])
        # print("data:\n%s" %(data[:].decode("utf16").replace("\r", "\r\n")))
        res += data[:].decode("utf16").replace("\r", "\r\n")
    return res


if __name__ == "__main__":
    file_name = sys.argv[1]
    # print("file_name", file_name)
    if olefile.isOleFile(file_name):
        res = parse_doc(file_name)
        print(res)
    else:
        print("zip")
        res = parse_docx(file_name)
        # print(res, type(res), res.decode('utf8'))
        print(res)