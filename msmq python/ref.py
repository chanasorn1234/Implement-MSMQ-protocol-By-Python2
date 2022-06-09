# from struct import *
# x = unpack('>ddddi', '\x00\x00\x00'*12)
# print(x)
import struct
# result = int(val.encode('bytes'), 16)
# result = int(val.encode('bytes'), 16)

# x=bytes(0)
# print(type(x))

# s = '\x00\x00\x00\x01\x00\x00\x00\xff\xff\x00\x00'
# print(type(s))
# print(struct.unpack('11B',s))

a = bytearray([1,2,3])
print(a)