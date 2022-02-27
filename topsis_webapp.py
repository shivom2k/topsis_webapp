import pandas as pd
import streamlit as st
import sys
import logging
import numpy as np
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import os 
def send_mail(email_id,resultfile):
    fromaddr = "shivuchawla2@gmail.com"
    toaddr = email_id
   
# instance of MIMEMultipart
    msg = MIMEMultipart()
  
# storing the senders email address  
    msg['From'] = fromaddr
  
# storing the receivers email address 
    msg['To'] = toaddr
  
# storing the subject 
    msg['Subject'] = "TOPSIS SCORE AND RANK GENERATOR" 
  
# string to store the body of the mail
    body = '''For the given input file(.csv), here is your ouput(.csv) file with topsis score and rank information provided for MCDM(Multiple Criteria Decision Making'''
  
# attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))
  
# open the file to be sent 
    filename = resultfile
    attachment = open(resultfile, "rb")
  
# instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')
  
# To change the payload into encoded form
    p.set_payload((attachment).read())
  
# encode into base64
    encoders.encode_base64(p)
   
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
  
# attach the instance 'p' to instance 'msg'
    msg.attach(p)
  
# creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)
  
# start TLS for security
    s.starttls()
  
# Authentication
    s.login(fromaddr, "12345678shivam")
  
# Converts the Multipart msg into a string
    text = msg.as_string()
  
# sending the mail
    s.sendmail(fromaddr, toaddr, text)
  
# terminating the session
    s.quit()
                    
    
def topsis(dataset,weight,impact,result_file_name,email_id):
    logging.basicConfig(filename=result_file_name+"-log.txt", level=logging.INFO)
    df=dataset
    model_name=list(df[df.columns[0]])
    data=df.iloc[:,1:]
    total_columns=data.shape[1]
    total_rows=data.shape[0]
    if total_columns<=2:
            logging.error('enter more than 3 columns in file')
            raise Exception("nter more than 3 columns in file")
            logging.shutdown() 
    for col in df.columns:
            if df[col].isnull().values.any():
                logging.error(f"{col} contains null values")
                raise Exception(f"{col} contains null values")
                logging.shutdown()         
    for i in range(total_rows):
        for j in range(total_columns):
            if np.char.isnumeric(str(data.iloc[i,j]))==True:
                logging.error('number are not numeric')
                raise Exception("number are not numeric")
                logging.shutdown() 
                
    weight=weight.split(",")
    #print(weight)
    impact=impact.split(",")
    if(len(weight)!=total_columns):
            logging.error('enter weights properly')
            raise Exception("enter weight properly")
            logging.shutdown() 
        
    if(len(impact)!=total_columns):
            logging.error('enter impact properly')
            raise Exception("enter impact properly")
            logging.shutdown() 
    
    for i in impact:
        if i!="+" and i!="-":
            logging.error('impact must be positive or negative')
            raise Exception("impact must be positive or negative")
            logging.shutdown() 
            
    
    for i in range(total_columns):
        temp=0;
        for j in range(total_rows):
            temp=temp+data.iloc[j,i]**2
        temp=temp**0.5
        for j in range(total_rows):
            #print(data.iloc[j,i])
            data.iat[j,i]=(data.iloc[j,i]/temp)*float(weight[i])

    vj_positive=[]
    vj_negative=[]
    for i in range(total_columns):
        if impact[i]=="+":
            vj_positive.append(data.iloc[:,i].max())
            vj_negative.append(data.iloc[:,i].min())
        if impact[i]=="-":
            vj_positive.append(data.iloc[:,i].min())
            vj_negative.append(data.iloc[:,i].max())
    sj_positive=[]
    sj_negative=[]
    for i in range(total_rows):
        temp=0
        temp1=0
        for j in range(total_columns):
            temp=temp+(vj_positive[j]-data.iloc[i,j])**2
            temp1=temp1+(vj_negative[j]-data.iloc[i,j])**2
        sj_positive.append(temp**0.5)
        sj_negative.append(temp1**0.5)
    topsis_score=[]
    for i in range(len(sj_positive)):
        a=sj_negative[i]/(sj_negative[i]+sj_positive[i])
        topsis_score.append(a)
    arr=np.array(topsis_score)
    index=np.argsort(arr)[::-1]

    df["topsis_score"]=topsis_score
    topsis_score = pd.DataFrame(topsis_score)
    topsis_rank = topsis_score.rank(method='first',ascending=False)
    df["rank"]=topsis_rank
    df = df.astype({"rank": int})
    df.to_csv(result_file_name,index=False)
    send_mail(email_id,result_file_name)
    st.success("Check your email, result file is successfully sent")

st.title("Topsis Calculator by Shivom")
if __name__ == '__main__':
    uploaded_file = st.file_uploader("Choose a file")
    user_input = st.text_input("enter weight seperated by comma")
    user_input1 = st.text_input("enter impact seperated by comma")
    user_input2 = st.text_input("enter the email")
    submit_button=st.button("Send")
    if submit_button:
                    try:
                        # if uploaded_file.name.split('.')[1]!="csv" and uploaded_file.split('.')[1]!="xlsx":
                        #     st.error("File format must be a csv of excel file")
                        # if uploaded_file.name.split('.')[1]=="csv":
                        #     spectra_df = pd.read_csv(uploaded_file)
                        # elif uploaded_file.name.split('.')[1]=="xlsx":
                        #     read = pd.read_excel (uploaded_file)
                        #     read.to_csv ('file_new', index = None,header=True)
                        spectra_df = pd.read_csv(uploaded_file)
                        topsis(spectra_df,user_input,user_input1,"result_topsis.csv",user_input2)
                    except Exception as e:
                        st.error(e)    