# coding: utf-8

# In[124]:


import pandas as pd


# In[125]:
print("Welcome to Automated Incident Metrics")
excel = input("Enter the path of your Incident excel file, eg: D:\Shubham\incident.xlsx :")
#data = pd.read_excel(r"C:/Users/SN5034547/Downloads/incident (25).xlsx")
data = pd.read_excel(r"{}".format(excel))


# In[126]:


total_tickets = data['Severity'].count()
#uniq = data['Severity'].unique()
slas = data["SLA Breached"].value_counts()
sla_missed = slas[True]
sla_miss=(sla_missed/total_tickets ) * 100


# In[127]:


sevs = data['Severity'].value_counts()

try:
    sev1 = sevs['Sev1']
except KeyError:
    sev1=0
try:
    sev2 = sevs['Sev2']
except KeyError:
    sev2=0
try:
    sev3 = sevs['Sev3']
except KeyError:
    sev3=0
try:
    sev4 = sevs['Sev4']
except KeyError:
    sev4=0
try:
    sev5 = sevs['Sev5']
except KeyError:
    sev5=0
try:
    sev6 = sevs['Sev6']
except KeyError:
    sev6=0


# In[128]:


#print(sev1,sev2,sev3,sev4,sev5,sev6)


# In[129]:


SERVER = "XXXX.XXXXXorg.com"
FROM = input("Enter Sender email address:")
name = FROM.split('_')[0]
#FROM = "Shubham_Naik@Syntelinc.com"
TO = ["Shubham_Naik@Syntelinc.com"] # must be a list
CC = ["Shubham_Naik@Syntelinc.com"] # must be a list
BCC=[]
#BCC = ["Nikhil_Jain@Syntelinc.com"] # must be a list
TO_ADD = [TO] + CC + BCC

# In[130]:


SUBJECT = "Incident metrics for PSR data"
TEXT = '''
Hi Arbazuddin,
1.Total Ticket: {}
2.SLA Misses: {:.2f} %
Incident Metrics 
o    Number of tickets per person Offshore: 2
o    Number of tickets per person Onsite: 1   
o    Number of hours per ticket Offshore: 2
o    Number of hours per ticket Onsite:  1   
o    Number of breached INCs : {}
o    Sev 1 : {}
o    Sev 2 : {}
o    Sev 3 : {}    
o    Sev 4 : {}
o    Sev 5 : {}
o    Sev 6 : {}
Thanks,
{}|GMS
'''.format(total_tickets,sla_miss,sla_missed,sev1,sev2,sev3,sev4,sev5,sev6,name)


# In[131]:


# Prepare actual message
message = "From: %s\r\n" % FROM + "To: %s\r\n" % ",".join(TO) + "CC: %s\r\n" % ",".join(CC) + "BCC: %s\r\n" % ",".join(BCC) + "Subject: %s\r\n" % SUBJECT + "\r\n"  + TEXT


# In[132]:


# Send the mail
import smtplib
server = smtplib.SMTP(SERVER,587)
server.sendmail(FROM, TO,message)
server.quit()


