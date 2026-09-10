import numpy as np
import timeit, time
import scipy

arr1 = np.random.random(100)  # 100k random values
arr2 = np.random.random(100)  # 100k random values

timeit
print(np.convolve(arr1, arr2))

timeit
print(scipy.signal.fftconvolve(arr1, arr2))

