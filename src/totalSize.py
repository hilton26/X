#! python3

# Lesson 30

import os

totalSize = 0
for filename in os.listdir('C:\\'):
    if not os.path.isfile(os.path.join('C:\\',filename)):
                          continue
    totalSize = totalSize + os.path.getsize(os.path.join('C:\\',filename))
print(totalSize)
