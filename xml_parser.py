import pandas as pd
import xml.etree.ElementTree as ET
import re, os,datetime
import graphviz as gv
from numpy import nan

path=r'C:\Users\e650188\Desktop\automation_files\xml_part_c'
mart_name='Part_C'
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

print('Parse xml')################## Parse xml #################
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

print('GRAPH')################## GRAPH #################
os.environ["PATH"] += os.pathsep + r'C:\Users\e650188\AppData\Local\graphviz238\bin'

dot = gv.Digraph(node_attr={'shape': 'box'},edge_attr={'dir':'forward'},graph_attr={'rankdir':'RL'})

for num in range(df.shape[0]):
    dot.edge(df.iloc[num]['ORG']+'  '+str(df.iloc[num]['R_num']),df.iloc[num]['name']+'  '+str(df.iloc[num]['level']))

dot.render(r'C:\Users\e650188\Desktop\automation_files\test-output\{}_{}.gv'.format(mart_name,(str(datetime.datetime.now())[:-7]).replace(":","-")), view=True)

print('DM formula')################### DM formula #################
dfcol_param=['DM_Name','Parameter','Default_Value'] # if this list changes often, will need a dict
df_param=pd.DataFrame(columns=dfcol_param)

dfcol_element = ['DM_Name','Element','status','alias','key','Formula','displayName']  # if this list changes often, will need a dict
df_element=pd.DataFrame(columns=dfcol_element)

dfcol_join = ['DM_Name','alias','Source_Element','Target_Element','Main_Table(Target)']
df_join=pd.DataFrame(columns=dfcol_join)

dfcol_filter=['DM_Name','Filter'] # if this list changes often, will need a dict
df_filter=pd.DataFrame(columns=dfcol_filter)

for r,d,files in os.walk(path):
    for file in files:
        if '.xml' in file:  # only opens up xml files
            with open(os.path.join(r,file), 'r') as xml_file: # do not decode utf-8
                etree = ET.parse(xml_file)

            
            for DataMart in etree.getroot():  # DataMart
                for item in DataMart.findall('Parameters/Parameter'):
                    df_param=df_param.append(pd.Series([DataMart.get('name'),
                          item.get('name'),
                          (None if item.find('DefaultValue') is None else item.find('DefaultValue').text)
                          ],index=dfcol_param),ignore_index=True)

                for item in DataMart.findall('Elements/Element'): # --element level
                    df_element=df_element.append(pd.Series([DataMart.get('name'),
                          item.get('name'),
                          item.get('status'),
                          item.get('alias'),
                          item.get('key'),
                          (None if item.find('Formula') is None else item.find('Formula').text),
                          item.get('displayName')
                          ],index=dfcol_element),ignore_index=True)
    
                for item in DataMart.findall('FilterCondition'):  # Filter
                    df_filter=df_filter.append(pd.Series([DataMart.get('name'),
                          (None if item.find('FilterExpression') is None else item.find('FilterExpression').text)
                          ],index=dfcol_filter),ignore_index=True)
                    
                for item in DataMart.findall('JoinedTable'):  # Joined Table
                    df_join=df_join.append(pd.Series([DataMart.get('name'),
                          item.get('alias'),
                          None if item.find('JoinElement') is None else item.find('JoinElement').get('sourceElementAlias'),
                          None if item.find('JoinElement') is None else item.find('JoinElement').get('targetElementAlias'),
                          None if item.find('JoinElement') is None else item.find('JoinElement').get('targetTable')
                          ],index=dfcol_join),ignore_index=True)

writer = pd.ExcelWriter(r'C:\Users\e650188\Desktop\automation_files\test-output\{}_output.xlsx'.format(mart_name))
df_element.to_excel(writer,'Elements',index=False)
df_join.to_excel(writer,'JoinTable',index=False)
df_filter.to_excel(writer,'Filter',index=False)
df_param.to_excel(writer,'Parameters',index=False)
writer.save()
