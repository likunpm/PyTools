import json
import pandas as pd
import datetime

file="/Users/likun/Library/Application Support/Google/Chrome/Default/Bookmarks"
fileto='/Users/likun/Desktop/bookmark.xlsx'

#谷歌时间戳转化
def getFiletime(dt):
    dt=hex(dt*10)[2:17]
    microseconds = int(dt, 16) / 10
    seconds, microseconds = divmod(microseconds, 1000000)
    days, seconds = divmod(seconds, 86400)

    ldt=datetime.datetime(1601, 1, 1) + datetime.timedelta(days, seconds, microseconds)
    return format(ldt,'%D')
    
    
class GoogleShuqian():
    def __init__(self,file):
        self.file=file
        self.res={}
    def GetShuqian(self):
        ShuQian=json.load(open(file)) #获取书签
        sq=pd.DataFrame(ShuQian['roots']['bookmark_bar']['children'])
        
        NoCate=sq[sq['type']=='url'] #没有分类的
        NoCate['date_added']=NoCate['date_added'].apply(lambda x:getFiletime(int(x)))
        NoCate=NoCate[['date_added','name','url']]
        self.res['没有分类']=NoCate
        
        Cate=sq[sq['type']=='folder'] #有分类的
        for i in Cate.values:
            CateName=i[4]
            CateList=pd.DataFrame(i[7])[['date_added','name','url']]
            CateList['date_added']=CateList['date_added'].apply(lambda x:getFiletime(int(x)))
            self.res[CateName]=CateList
        return self.res 
        
        
def WriteToxls():
   writer=pd.ExcelWriter(fileto) 
   for key,values in GoogleShuqian(file).GetShuqian().items():
       values.to_excel(writer,sheet_name=key,index=None) 
   writer.save()
   return 
   
if __name__ == '__main__':
    WriteToxls()
