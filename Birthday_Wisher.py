import pandas as pd
import datetime
import smtplib

GMAIL_ID='gogosingh06@gmail.com'
GMAIL_PASS='gogo@234'

def sendmail(to,sub,msg):
    print(f"Email sent to {to} with subject {sub} with message as {msg}")
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PASS)
    s.sendmail(GMAIL_ID,to,f"Subject: {sub}\n\n{msg}")
    s.quit()
if __name__=="__main__":
    # sendmail(GMAIL_ID,"subject","test")
    df=pd.read_excel("data.xlsx")
    # print(df)
    today=datetime.date.today().strftime("%d-%m")
    yearNow=datetime.date.today().strftime("%Y")
    writeInd=[]
    # print(today)
    for index,item in df.iterrows():
        # print(index,item["Birthday"])
        bday=item["Birthday"].strftime("%d-%m")
        # print(bday)
        if(today==bday)and yearNow not in str(item["Year"]):
            sendmail(item["Email"],"Happy Birthday",item["Message"])
            writeInd.append(index)
    # print(writeInd)
    if writeInd:
        for i in writeInd:
            yr = df.loc[i,'Year']
            df.loc[i,'Year']=f"{yr},{yearNow}"
        # print(df)
        df.to_excel("data.xlsx",index=False)