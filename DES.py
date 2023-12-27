from Cryptodome.Cipher import DES
from Cryptodome.Util.Padding import pad, unpad

def encryptDES(key, text):
    key = bytearray(key,'utf8')
    text = bytearray(text, 'utf8')
    print("---", key)
    print("+++", text)
    des = DES.new(key, DES.MODE_ECB)

    encrypted_text = des.encrypt(pad(text,64))

    Str = ''
    for i in encrypted_text:
        Str += '0'*(8-len(bin(i)[2:])) + bin(i)[2:]

    return Str


def bitstring_to_bytes(s):
    v = int(s, 2)
    b = bytearray()
    while v:
        b.append(v & 0xff)
        v >>= 8
    return bytes(b[::-1])

def bytes_to_bitstring(s):
    v = int(s, 2)
    b = bytearray()
    while v:
        b.append(v & 0xff)
        v >>= 8
    return bytes(b[::-1])


def decryptDES(key,Str):
    strde = bitstring_to_bytes(Str)
    key = bytearray(key,'utf8')
    des = DES.new(key, DES.MODE_ECB)

    return unpad(des.decrypt(strde),64).decode('utf-8')

