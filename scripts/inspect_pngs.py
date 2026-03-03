import os,struct
p=os.path.join(os.path.dirname(__file__), '..', 'image')
if not os.path.isdir(p):
    p=os.path.join(os.path.dirname(__file__), 'image')
files=sorted(os.listdir(p))
for f in files:
    path=os.path.join(p,f)
    try:
        with open(path,'rb') as fh:
            sig=fh.read(8)
            if sig!=b'\x89PNG\r\n\x1a\n':
                print(f+'\tNOT_PNG')
                continue
            texts=[]
            width=height=None
            while True:
                head=fh.read(8)
                if len(head)<8:
                    break
                length,ctype=struct.unpack('>I4s',head)
                data=fh.read(length)
                crc=fh.read(4)
                if ctype==b'IHDR':
                    width,height=struct.unpack('>II',data[:8])
                if ctype in (b'tEXt',b'iTXt'):
                    try:
                        texts.append(data.decode('latin1'))
                    except Exception:
                        try:
                            texts.append(data.decode('utf-8','ignore'))
                        except:
                            texts.append(repr(data))
                if ctype==b'IEND':
                    break
            print(f"{f}\t{width}x{height}\ttexts={texts}")
    except Exception as e:
        print(f+"\tERR\t"+str(e))
