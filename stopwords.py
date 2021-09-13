import sys

result=[]

with open('/Users/tyrzax/Downloads/stopwords-master/baidu_stopwords.txt','r') as f:
	for line in f:
		result.append(line)
f.close()
with open('/Users/tyrzax/Downloads/stopwords-master/cn_stopwords.txt','r') as a:
	for line in a:
		result.append(line)
f.close()
with open('/Users/tyrzax/Downloads/stopwords-master/scu_stopwords.txt','r') as b:
	for line in b:
		result.append(line)
f.close()
with open('/Users/tyrzax/Downloads/stopwords-master/hit_stopwords.txt','r') as c:
	for line in c:
		result.append(line)
f.close()

stopwords = []
for i in result:
    if not i in stopwords:
        stopwords.append(i)

with open('/Users/tyrzax/ZJU Documents/stopwords.txt','w') as e:
    for word in stopwords:
        e.writelines(word)
e.close()