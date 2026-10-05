import re
import sys
import io

path = sys.argv[1]
keep = sys.argv[2].split(",") if len(sys.argv) > 2 else []
text = io.open(path, encoding="utf-8", errors="replace").read()
text = re.sub(r"\x1b\[[0-9;]*m", "", text)
lines = text.splitlines()
for line in lines:
    if not keep or any(token in line for token in keep):
        print(line)
