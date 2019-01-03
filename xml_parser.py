import pandas as pd
import xml.etree.ElementTree as ET
import re, os,datetime
import graphviz as gv
from numpy import nan

path=r'C:\Users\e650188\Desktop\automation_files\xml - Copy'

################ Prepare xml ################

for r,d,files in os.walk(path):
    for file in files:
        if '.txt' in file:      # only opens up txt files
            print(file)
            with open(os.path.join(r,file), encoding="utf8",mode="r") as textobj:
                content = list(textobj)
            for num, txt in enumerate(content):
                if r'<?xml version=' in txt:
                    d=num # determine xml starts at which row
            for del_line in range(-(d-1),1): # delete from row 0 will make row 1 become row 0, deleting from the bottom will keep the correct order
                del content[-del_line]
            with open(os.path.join(r,file)[:re.search(".txt",os.path.join(r,file)).span()[0]]+"_{}.xml".format(re.findall(r'Data=.*name=\"([^\"]*)',content[1])[0]),"w") as textobj:
                for line in content:
                    textobj.write(line)
textobj.close()

################## Parse xml #################
def iter_docs(author,element='Table'):
    author_attr = author.attrib
    for doc in author.iter(element):
        doc_dict = author_attr.copy()
        doc_dict.update(doc.attrib)
        doc_dict['source'] = doc.text
        yield doc_dict
        
df=pd.DataFrame([])
for r,d,files in os.walk(path):
    for file in files:
        if '.xml' in file:  # only opens up xml files
            with open(os.path.join(r,file), 'r') as xml_file: # do not decode utf-8
                etree = ET.parse(xml_file)
            for element in ['Table','DataFeed','JoinedTable']:
                doc_df = pd.DataFrame(list(iter_docs(etree.getroot(),element)))
                if not doc_df.empty:
                    doc_df['ORG']=file[re.search('Response_',file).span()[1]:-4]
                    doc_df['R_num']=int(file[:re.search('Response_',file).span()[0]-1])
                    df=pd.concat([df,doc_df[['R_num','ORG','name']]],axis=0)
df=df.sort_values(by='R_num', ascending=True)
df=df.reset_index()
df=df.drop('index',axis=1)
df['level']=nan  # using nan to keep all value float


for num,nam in enumerate(df.name):
    df.loc[num,'level']=next((df.loc[i].R_num for i, val in enumerate(df.ORG) if (val == nam and i>num and (df.loc[i].R_num not in list(df.level)))), nan)
df.level=df.level.map('{:.0f}'.format)  # int format

df.loc[df['level']=='nan','level'] = df[df['level']=='nan']['R_num']  # replace nan with R_num to distinct feeds

################## GRAPH #################
os.environ["PATH"] += os.pathsep + r'C:\Users\e650188\AppData\Local\graphviz238\bin'

dot = gv.Digraph(node_attr={'shape': 'box'},edge_attr={'dir':'forward'},graph_attr={'rankdir':'RL'})

for num in range(df.shape[0]):
    dot.edge(df.iloc[num]['ORG']+'  '+str(df.iloc[num]['R_num']),df.iloc[num]['name']+'  '+str(df.iloc[num]['level']))

dot.render(r'C:\Users\e650188\Desktop\automation_files\test-output\round-table_{}.gv'.format(str(datetime.datetime.now())[:-7].replace(":","-")), view=True)

################## DM element #################

df1=pd.DataFrame([])
for r,d,files in os.walk(path):
    for file in files:
        if '.xml' in file:  # only opens up xml files
            with open(os.path.join(r,file), 'r') as xml_file: # do not decode utf-8
                etree = ET.parse(xml_file)
            for element in ['Element']:  #  TBC 'Formula','FilterExpression','JoinElement'
                doc_df1 = pd.DataFrame(list(iter_docs(etree.getroot(),element)))
                if not doc_df1.empty:
                    doc_df1['ESP_Table_Name']=file[re.search('Response_',file).span()[1]:-4]
                    doc_df1['R_num']=int(file[:re.search('Response_',file).span()[0]-1])
                    df1=pd.concat([df1,doc_df1],axis=0)

df1=df1[['status','ESP_Table_Name','alias','name','displayName','key','dataType','length','physicalTableName','physicalColumnName','rollupFunction','sourceAlias','lookupCategory','lookupElement','R_num']]

################### DM formula #################

def get_value_of_node(node):
    return node.text if node is not None else nan

df2=pd.DataFrame([])
for r,d,files in os.walk(path):
    for file in files:
        if '.xml' in file:  # only opens up xml files
            with open(os.path.join(r,file), 'r') as xml_file: # do not decode utf-8
                etree = ET.parse(xml_file)
            dfcols = ['Element', 'Formula', 'FilterExpression', 'JoinElement']
            df_xml=pd.DataFrame(columns=dfcols)
            
            for node in etree.getroot():
                Element=node.attrib.get('name')
                Formula=node.find('Formula')
                FilterExpression=node.find('FilterExpression')
                JoinElement=node.find('JoinElement')
                df_xml= df_xml.append(pd.Series([Element, get_value_of_node(Formula), get_value_of_node(FilterExpression),get_value_of_node(JoinElement)], index=dfcols),ignore_index=True)
