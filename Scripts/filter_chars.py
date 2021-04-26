#!/usr/local/opt/python3/bin/python3

import os
import glob
import re

top = "."
for (dirpath, dirnames, filenames) in os.walk(top):
    pattern = os.path.join(dirpath, "*.ps1")
    for path in glob.iglob(pattern):
        print(path)
        with open(path, "r") as cin:
            text = cin.read()
        text = re.sub(r"\|\s+", "| ", text)
        with open(path, "w") as cout:
            cout.write(text)
