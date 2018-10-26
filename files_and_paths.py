import os
from pprint import pprint

files = []
for dirname, dirnames, filenames in os.walk('.'):
    # print path to all subdirectories first.
    for subdirname in dirnames:
        files.append(os.path.join(dirname, subdirname))

    # print path to all filenames.
    for filename in filenames:
        files.append(os.path.join(dirname, filename))


pprint(files)


