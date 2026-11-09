import sys, numpy as np
from PIL import Image
from scipy.ndimage import gaussian_filter1d
from scipy.signal import find_peaks
f=sys.argv[1]
a=255-np.asarray(Image.open(f).convert('L')).astype(float)
for y0 in range(300,1100,100):
    c=gaussian_filter1d(a[y0:y0+60].mean(0),8)
    # lanes = local maxima; gaps = minima
    mins,_=find_peaks(-c,distance=80,prominence=5)
    print(y0, [int(m) for m in mins if 250<m<1800])
