import os

f1 = 'MonsterParamEx2_Copy.xml'
f2 = 'MonsterParamEx2.xml'

with open(f1, 'rb') as file1, open(f2, 'rb') as file2:
    c1 = file1.read()
    c2 = file2.read()

print(f"Size 1: {len(c1)}, Size 2: {len(c2)}")
print(f"BOM 1: {c1[:3]}, BOM 2: {c2[:3]}")
print(f"Line endings 1 CRLF: {c1.count(b'\\r\\n')}, LF: {c1.count(b'\\n')}")
print(f"Line endings 2 CRLF: {c2.count(b'\\r\\n')}, LF: {c2.count(b'\\n')}")
print(f"Ends with 1: {c1[-20:]}")
print(f"Ends with 2: {c2[-20:]}")
