import os
import shutil
import time

src = "platform"
dst = "webapp"

for i in range(5):
    try:
        shutil.move(src, dst)
        print(f"Successfully renamed {src} to {dst}")
        break
    except Exception as e:
        print(f"Attempt {i+1} failed: {e}")
        time.sleep(1)
