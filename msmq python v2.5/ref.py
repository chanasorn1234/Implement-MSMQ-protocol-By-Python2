# from struct import *
# x = unpack('>ddddi', '\x00\x00\x00'*12)
# print(x)
from time import sleep
import os
import logging
from logging.handlers import TimedRotatingFileHandler
# result = int(val.encode('bytes'), 16)
# result = int(val.encode('bytes'), 16)

# x=bytes(0)
# print(type(x))

# s = '\x00\x00\x00\x01\x00\x00\x00\xff\xff\x00\x00'
# print(type(s))
# print(struct.unpack('11B',s))

# a = bytearray([1,2,3])
# print(a)
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter(fmt=' %(asctime)s   [%(levelname)s]   |  %(message)s',
                                        datefmt='%m-%d-%y %H:%M:%S')
if(1):
    inp = input('Enter number:')
    print(inp)
    res_computer_name = os.getenv('COMPUTERNAME')
    print(res_computer_name)
    sleep(10)
    fh = TimedRotatingFileHandler('%s/log/' % (inp), when="midnight", interval=1, encoding='utf8')
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    logger.info('your input',inp)
else:
    print('program not read')
    logger.error('program not read')

