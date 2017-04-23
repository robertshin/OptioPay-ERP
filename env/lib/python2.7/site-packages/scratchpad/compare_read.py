import time
from memory_profiler import memory_usage
from pympler.muppy import print_summary
from xlrd import open_workbook
from openpyxl import load_workbook

fname = "Issues/bug494.xlsx"

def rd():
    start = time.clock()
    wb = open_workbook(fname)
    end = time.clock()
    print("xlrd {:0.1f}s".format(end - start))
    use = memory_usage(proc=-1, interval=1)[0]
    print("Memory use %.1f MB" % use)
    #print_summary()


def pixel():
    start = time.clock()
    wb = load_workbook(fname, keep_links=False)
    end = time.clock()
    print("openpyxl {:0.1f} s".format(end - start))
    use = memory_usage(proc=-1, interval=1)[0]
    print("Memory use %.1f MB" % use)
    #print_summary()


if __name__ == '__main__':
    from openpyxl import __version__
    print(__version__)
    rd()
    #pixel()
