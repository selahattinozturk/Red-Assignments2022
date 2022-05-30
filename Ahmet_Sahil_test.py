
import pandas as pd
def Ahmet_Sahil_test():
        
    ab=pd.read_excel("4- BOBİ FRS Nakit Akış Tablosu - Dolaylı Yöntem (Konsolide).xlsx")
    data=ab.fillna(0)
    new_df=pd.DataFrame(columns=['Açıklamalar', 'Cari Dönem', 'Önceki Dönem'])
    
    for i in range(2,71):
        if data.iloc[i,1]!=0:
            
            value1=data.iloc[i,1]
        elif data.iloc[i,2]!=0:
            
            value1=data.iloc[i,2]
    
        elif data.iloc[i,3]!=0:
            
            value1=data.iloc[i,3]
    
        elif data.iloc[i,4]!=0:
            
            value1=data.iloc[i,4]
        else:
           
            continue
        new_df.loc[len(new_df.index)] = [value1, 0, 0]
    return new_df
DataBase=Ahmet_Sahil_test()
