import pandas as pd

df = pd.read_csv('C:/Users/11359023/Desktop/engmix.txt', sep=" ", header=None,encoding = "latin-1")
df.columns = ["dictionary_text"]

path = "C:/Users/11359023/Desktop/exam/"

for i in df["dictionary_text"]:
#Create and Write
    try:
        if len(i)<=1:
            f = open(path+str(i[0].upper())+"/_"+str(i[0].upper())+"/"+str(i)+".txt", "w")
            for j in range(1,101):
                f.write(str(i)+"\n\
")
            f.close()
        else:
            f = open(path+str(i[:1].upper())+"/_"+str(i[1:2].upper())+"/"+str(i)+".txt", "w")
            for j in range(1,101):
                f.write(str(i)+"\n\
")
            f.close()
    except:
        pass
