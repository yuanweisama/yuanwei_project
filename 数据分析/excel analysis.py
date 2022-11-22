import numpy
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
matplotlib.use('TkAgg')


'''path_excel_name1='mars_tianchi_songs.csv'
path_excel_name = 'mars_tianchi_user_actions.csv'
data = pd.read_csv(path_excel_name)
columns = data.columns
x = data.iloc[:1000, 3]
# distname=open('(1000)mars_tianchi_songs.csv','ab')
# pd_10=pd.DataFrame(x)#将列表转换为DataFrame格式
# pd_10.to_csv(distname,index=False)
y = data.iloc[:1000, 3]
# print(type(x))
# print(type(y))
plt.hist(x, 50)
plt.scatter(y,x)
plt.show()
'''

path_excel_name1='(1000)mars_tianchi_songs.csv'
path_excel_name = '(1000)mars_tianchi_user_actions.csv'
data = pd.read_csv(path_excel_name)
columns = data.columns
x = data.iloc[:1000, 2]
y = data.iloc[:1000, 3]
print(x)
print(y)
# 画第1个图：折线图
plt.subplot(221)
plt.plot(x,y)
# 画第2个图：散点图
plt.subplot(222)
plt.scatter(x,y)
# 画第3个图：条形图
# 前面的两个图占了221和222的位置，如果想在下面只放一个图，得把前两个当成一列，即2行1列第2个位置
plt.subplot(212)
plt.bar(x,10)
plt.show()