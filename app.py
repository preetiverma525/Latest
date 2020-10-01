#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import flask
from flask import request, jsonify


app = flask.Flask(__name__)
app.config["DEBUG"] = True

# Create some test data for our catalog in the form of a list of dictionaries.

@app.route('/', methods=['GET'])
def home():
    return '''<h1>WELCOME</h1>
<p>Get all pppp hhhh insterted and delected text from the word document.</p>'''

@app.route('/api', methods=['GET'])


def api_id():
    import requests
    import xmltodict
    import pprint
    import json
    from zipfile import ZipFile
    from urllib.request import urlopen
    from io import BytesIO
#     from zipfile import *
    from bs4 import BeautifulSoup
    import pandas as pd
    
    if 'Id' in request.args:
        Id = request.args['Id']
        print(type(Id))
    else:
        return "Error: No id field provided. Please specify an id."
    results = []
    url='https://bps.quickbase.com/db/bqsuve7qy?a=API_GetRecordInfo&rid='+Id+'&fmt=structured&usertoken=b5d3bm_kn5b_d4c4hh5b76ajm4t6fzg96bbber'
    #url='https://builderprogram-pverma.quickbase.com/db/bqscz87a5?a=API_GetRecordInfo&rid='+Id+'&fmt=structured&usertoken=b5fdma_nx3z_pzc2d4b2uvnihbayfyd8bk8swsk'
    response = requests.request("GET", url)
    print(response)
    r=response.text.encode('utf8')
    print(r)
    pp = pprint.PrettyPrinter(indent=4)
    data=json.dumps(xmltodict.parse(r))
    data1=json.loads(data)
    for i in data1.values():
        for j in i['field']:
            if j['fid']=="11":
                Document=j['value']
            if j['fid']=="3":
                R_Id=j['value']
        Mapped_data={R_Id :Document}
#         print(Mapped_data)
#         return Mapped_data
    track_changed_for_del=[]
    track_changed_for_ins=[]
    for key, value in Mapped_data.items():
        wordfile=urlopen(value).read()
        wordfile=BytesIO(wordfile)
        document=ZipFile(wordfile)
        document.namelist()
        xml_content=document.read('word/document.xml')
        wordobj=BeautifulSoup(xml_content.decode('utf-8'),'xml')
        key_record=key
        for dl in wordobj.find_all('w:del'):
            Text=dl.text
            author=dl.get('w:author')
            Date=dl.get('w:date')
            Type='Deleted Text'
            ID=dl.get('w:id')
            ID=int(ID)
            dataDict_del = { 'Text':Text,'Author':author,'Date':Date,'Type':Type,'ID':ID,'Record_Id':key_record}
            print(dataDict_del)
            track_changed_for_del.append(dataDict_del)
        for ins in wordobj.find_all('w:ins'):
            Text=ins.text
            author=ins.get('w:author')
            Date=ins.get('w:date')
            Type='Inserted Text'
            ID=ins.get('w:id')
            ID=int(ID)
            dataDict_ins = { 'Text':Text,'Author':author,'Date':Date,'Type':Type,'ID':ID,'Record_Id':key_record}
            track_changed_for_ins.append(dataDict_ins)
        df_track_changed_ins= pd.DataFrame(track_changed_for_ins)
        df_track_changed_del= pd.DataFrame(track_changed_for_del)
        df_track_changed_del["Text"]= df_track_changed_del["Text"].replace(' ', "NaN")
        df_track_changed_ins["Text"]= df_track_changed_ins["Text"].replace('', "NaN")
        df_track_changed_del.drop(df_track_changed_del.loc[df_track_changed_del['Text']=='NaN'].index, inplace=True)
        df_track_changed_ins=df_track_changed_ins.sort_values(by='ID')
        df_track_changed_del=df_track_changed_del.sort_values(by='ID')
        all_data=df_track_changed_ins.append(df_track_changed_del)
        all_data=all_data.sort_values(by='ID',ascending=True)
        val=all_data.to_dict('records')
        csv_data=all_data.to_csv(index=False,header=False)
        usertoken = 'b5d3bm_kn5b_d4c4hh5b76ajm4t6fzg96bbber'
        dbid = 'bqi9wpaun'
        field_list='12.6.8.7.11.13'
        headers= {'Content-Type': 'application/xml', 'QUICKBASE-ACTION': 'api_importfromcsv'}
        data = '<qdbapi><records_csv><![CDATA[' + csv_data + ']]></records_csv><clist>'+field_list+'</clist>        <usertoken>' + usertoken + '</usertoken></qdbapi>'
        import requests
        r2 = requests.post(url='https://bps.quickbase.com/db/' + dbid, data=data.encode('utf-8'), headers=headers)
    return r2.text
#     else:
#         return "Error: No id field provided. Please specify an id."
#     for d in data:
#         if d['Record_Id'] == Id:
#             results.append(d)
#     return data1
    

app.run(debug=True,use_reloader=False)


# In[ ]:





# In[ ]:




