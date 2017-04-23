import time
from xml.etree.ElementTree import parse, iterparse
import os

filename = "../xml/xl/worksheets/sheet4.xml"


def iterative():
    start = time.clock()
    stream = iterparse(filename)
    for _, element in stream:
        pass
        #element.clear()
    stop = time.clock()
    print("Iterating took {0}s".format(stop))


def single_pass():
    start = time.clock()
    tree = parse(filename)
    stop = time.clock()
    print("Single pass took {0}s".format(stop))


if __name__ == '__main__':
    iterative()
    single_pass()


"""
Iterating took 31.690988s
Single pass took 48.358831s
"""
