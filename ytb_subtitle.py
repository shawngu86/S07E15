import re

s=open('STOP.txt',"r")
txt=s.read()

p=re.compile(r'</font>|<font color="#\S\S\S\S\S\S">|\d*:\d*:\d*,\d*\s-->\s\d*:\d*:\d*,\d*|\n')
s=p.split(txt)
while "" in s:
    s.remove("")
print(s)
