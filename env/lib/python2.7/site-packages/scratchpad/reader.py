from openpyxl import load_workbook, __version__
import time
import cProfile

from memory_profiler import memory_usage


def main():
    start = time.clock()
    wb = load_workbook("Issues/bug494.xlsx", keep_links=False)
    ws = wb.active
    print("openpyxl {0}".format(__version__))
    print(time.clock() - start)

if __name__ == '__main__':
    cProfile.run("main()", sort="tottime")
    from pympler.muppy import print_summary
    print_summary()
