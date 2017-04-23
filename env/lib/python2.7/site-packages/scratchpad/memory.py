import time
from openpyxl.styles.styleable import StyleArray

from memory_profiler import memory_usage


ids = []

start = time.clock()
for i in range(1000000):
    if i % 100000 == 0:
        print(i)
    ids.append(StyleArray())

use = memory_usage(proc=-1, interval=1)[0]
print("Memory use %s" % use)
print("Time = %d seconds" % (time.clock() - start))
