import sys, numpy as np
from PIL import Image
from scipy.ndimage import gaussian_filter1d
from scipy.signal import find_peaks
f,x0,x1=sys.argv[1],int(sys.argv[2]),int(sys.argv[3])
a=255-np.asarray(Image.open(f).convert('L')).astype(float)
p=gaussian_filter1d(a[:,x0:x1].mean(1),2)
pk,pr=find_peaks(p,prominence=4,distance=10)
for k in pk: print(k, round(p[k],1), round(pr['prominences'][list(pk).index(k)],1))
