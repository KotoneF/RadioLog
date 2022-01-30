# -*- coding: utf-8 -*-

import PySimpleGUI as SG
import re
import datetime
import copy
import xml.etree.ElementTree as ET
from xml.dom import minidom


SG.theme('DarkAmber')

#GUI用変数
#fontstyle = 'Consolas'
fontstyle = 'ＭＳ ゴシック'
fontsize = 10
bandList = ('160m','80m','75m','40m','30m','20m','17m','15m','12m','10m','6m','2m','70cm','23cm','13cm','5cm')
modeList = ('CW','SSB','AM','FM','FT8','C4FM','RTTY','SSTV')

# QSOデータ入力項目リスト
inputList = ('CALL','QSO_DATE','TIME_ON','TIME_OFF','RST_RCVD','RST_SENT','BAND','FREQ','MODE','COMMENT','NAME','JCC','GL','QTH','RIG','RX_PWR','ANTENNA','QSL_VIA','QSL_SENT','QSL_RCVD','MY_QTH','MY_RIG','TX_PWR','MY_ANTENNA','CONTEST_ID','SRX','STX')
# GUIレイアウト
layout = [
    [SG.Text('', key='tex_top',font=(fontstyle,fontsize))],
    [SG.FileBrowse('Open ADX',font=(fontstyle,fontsize),key='adxPath',target='tex_adx'),
     SG.InputText(key='tex_adx',font=(fontstyle,fontsize),size=(34,1), enable_events=True, readonly=True)],
    [SG.Text('Number',size=(7,1),font=(fontstyle,fontsize),enable_events=True),SG.Input(key='Number',size=(6,1),font=(fontstyle,fontsize)),
     SG.Button('Prev',font=(fontstyle,fontsize)),SG.Button('Next',font=(fontstyle,fontsize)),
     SG.Text('',size=(7,1),key='num_range',font=(fontstyle,fontsize))],
    [SG.Text('Required Information',font=(fontstyle,fontsize))],
    [SG.Text('C/S',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='CALL',size=(10,1),font=(fontstyle,fontsize))],
    [SG.Text('Date',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='QSO_DATE',size=(8,1),font=(fontstyle,fontsize)),
     SG.Text('TimeON',size=(7,1),font=(fontstyle,fontsize)),SG.Input(key='TIME_ON',size=(4,1),font=(fontstyle,fontsize)),
     SG.Text('TimeOFF',size=(7,1),font=(fontstyle,fontsize)),SG.Input(key='TIME_OFF',size=(4,1),font=(fontstyle,fontsize))],
    [SG.Text('RST Rcvd',size=(8,1),font=(fontstyle,fontsize)),SG.Input(key='RST_RCVD',size=(10,1),font=(fontstyle,fontsize)),
     SG.Text('RST Sent',size=(8,1),font=(fontstyle,fontsize)),SG.Input(key='RST_SENT',size=(10,1),font=(fontstyle,fontsize))],
    [SG.Text('Band',size=(4,1),font=(fontstyle,fontsize)),SG.Combo(bandList,key='BAND',size=(6,1),font=(fontstyle,fontsize)),
     SG.Text('Freq',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='FREQ',size=(9,1),font=(fontstyle,fontsize)),
     SG.Text('Mode',size=(4,1),font=(fontstyle,fontsize)),SG.Combo(modeList,key='MODE',size=(6,1),font=(fontstyle,fontsize))],
    [SG.Text('RMKS',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='COMMENT',size=(40,1),font=(fontstyle,fontsize))],
    [SG.Text('Partner Information',font=(fontstyle,fontsize))],
    [SG.Text('Name',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='NAME',size=(10,1),font=(fontstyle,fontsize)),
     SG.Text('JCC',size=(3,1),font=(fontstyle,fontsize)),SG.Input(key='JCC',size=(10,1),font=(fontstyle,fontsize)),
     SG.Text('GL',size=(3,1),font=(fontstyle,fontsize)),SG.Input(key='GL',size=(6,1),font=(fontstyle,fontsize))],
    [SG.Text('QTH',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='QTH',size=(40,1),font=(fontstyle,fontsize))],
    [SG.Text('Rig',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='RIG',size=(25,1),font=(fontstyle,fontsize)),
     SG.Text('PWR',size=(3,1),font=(fontstyle,fontsize)),SG.Input(key='RX_PWR',size=(4,1),font=(fontstyle,fontsize))],
    [SG.Text('Ant',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='ANTENNA',size=(40,1),font=(fontstyle,fontsize))],
    [SG.Text('QSL Card Information',font=(fontstyle,fontsize))],
    [SG.Text('QSL Via',size=(7,1),font=(fontstyle,fontsize)),SG.Input(key='QSL_VIA',size=(6,1),font=(fontstyle,fontsize)),
     SG.Text('QSL Sent',size=(8,1),font=(fontstyle,fontsize)),SG.Input(key='QSL_SENT',size=(1,1),font=(fontstyle,fontsize)),
     SG.Text('QSL Rcvd',size=(8,1),font=(fontstyle,fontsize)),SG.Input(key='QSL_RCVD',size=(1,1),font=(fontstyle,fontsize))],
    [SG.Text('My Information',font=(fontstyle,fontsize))],
    [SG.Text('QTH',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='MY_QTH',size=(40,1),font=(fontstyle,fontsize))],
    [SG.Text('Rig',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='MY_RIG',size=(25,1),font=(fontstyle,fontsize)),
     SG.Text('PWR',size=(3,1),font=(fontstyle,fontsize)),SG.Input(key='TX_PWR',size=(4,1),font=(fontstyle,fontsize))],
    [SG.Text('Ant',size=(4,1),font=(fontstyle,fontsize)),SG.Input(key='MY_ANTENNA',size=(40,1),font=(fontstyle,fontsize))],
    [SG.Text('Contest Information',font=(fontstyle,fontsize))],
    [SG.Text('Contest',size=(7,1),font=(fontstyle,fontsize)),SG.Input(key='CONTEST_ID',size=(37,1),font=(fontstyle,fontsize))],
    [SG.Text('Num RX',size=(6,1),font=(fontstyle,fontsize)),SG.Input(key='SRX',size=(10,1),font=(fontstyle,fontsize)),
     SG.Text('Num TX',size=(6,1),font=(fontstyle,fontsize)),SG.Input(key='STX',size=(10,1),font=(fontstyle,fontsize))],
    [SG.Button('Add',font=(fontstyle,fontsize)),
     SG.Button('Update',font=(fontstyle,fontsize)),
     SG.Button('Export',font=(fontstyle,fontsize)),
     SG.Button('JIS',font=(fontstyle,fontsize))]
]

# ADIF読み込みデータ変数
xmldata = ET.Element('ADX')
show_number = 1
is_new_qso = False

# 他変数
version = '0.2'

# Main
def main():
    global show_number, xmldata
    window = SG.Window('Radio Log ' + 'v.' + version, layout, return_keyboard_events=True)

    init_xml()

    while True:
        event, values = window.read()

        try:
            code = ord(event[0])
        except TypeError:
            code = 0

        # フォーカスしているオブジェクトのキーを取得
        cur_focus = window.find_element_with_focus()
        if cur_focus is not None:
            cur_focus = cur_focus.Key
        else:
            cur_focus = ""

        if event == SG.WIN_CLOSED :
            break

        # ADXファイル読み込みボタン押下
        elif event == 'tex_adx':
            if values['adxPath'] != '':
                adxImport(values['adxPath'])
                if len(xmldata.findall('./RECORDS/RECORD')) > 0 :
                    show_number = 1
                    window.find_element('Number').update(show_number)
                    input_form(xmldata.findall('./RECORDS/RECORD')[show_number - 1], window)
        # Prevボタン押下
        elif event == 'Prev':
            updateview(window, show_number - 1)
        # Nextボタン押下
        elif event == 'Next':
            updateview(window, show_number + 1)
        # Number上でEnterを押下
        elif code == 13 and cur_focus == 'Number':
            try:
                updateview(window, int(values['Number']))
            except ValueError:
                pass    
        # Addボタン押下
        elif event == 'Add':
            data_add(window)
            updateview(window, len(xmldata.findall('./RECORDS/RECORD')))
        # Updateボタン押下
        elif event == 'Update':
            data_update(window)
        # Exportボタン押下
        elif event == 'Export':
            adxExport('')
        # JIS-Exportボタン押下
        elif event == 'JIS':
            adxJisExport('')
    
        # データ範囲表記更新
        if len(xmldata.findall('./RECORDS/RECORD')) > 0:
            window.find_element('num_range').update('1 - ' + str(len(xmldata.findall('./RECORDS/RECORD'))))
        else:
            window.find_element('num_range').update('')

    window.close()


# XMLデータ初期化
def init_xml():
    global xmldata
    xmldata = ET.Element('ADX')
    header = ET.SubElement(xmldata, 'HEADER')
    timestamp = ET.SubElement(header, 'CREATED_TIMESTAMP')
    timestamp.text = datetime.datetime.now().strftime('%Y%m%d %H%M%S')
    adif_ver = ET.SubElement(header, 'ADIF_VER')
    adif_ver.text = '3.1.0'
    programid = ET.SubElement(header, 'PROGRAMID')
    programid.text = 'RadioLog'
    programversion = ET.SubElement(header, 'PROGRAMVERSION')
    programversion.text = version
    records = ET.SubElement(xmldata, 'RECORDS')

# Prevボタン押下処理
def updateview(window, num):
    global xmldata, show_number
    if num >= 1 and num <= len(xmldata.findall('./RECORDS/RECORD')):
        show_number = num
        window.find_element('Number').update(num)
        input_form(xmldata.findall('./RECORDS/RECORD')[num - 1], window)

# フォーム値代入
def input_form(record, window):
    # フォーム値クリア
    for k in inputList:
        window.find_element(k).update('')

    # データ代入
    dt = ''
    ton = ''
    toff = ''
    for d in record:
        window.find_element(d.tag).update(value=d.text) 
        if d.tag == 'QSO_DATE': dt = d.text
        elif d.tag == 'TIME_ON': ton = d.text
        elif d.tag == 'TIME_OFF': toff = d.text

    # 時間データをUTCからJSTへ変換
    if ton :
        ton_utc = datetime.datetime.strptime(dt+ton+'+0000', '%Y%m%d%H%M%z')
        ton_jst = ton_utc.astimezone(datetime.timezone(datetime.timedelta(hours=+9)))
        date_str = datetime.datetime.strftime(ton_jst,'%Y%m%d')
        ton_str = datetime.datetime.strftime(ton_jst,'%H%M')
        window.find_element('QSO_DATE').update(value=date_str)
        window.find_element('TIME_ON').update(value=ton_str)
    if toff :
        toff_utc = datetime.datetime.strptime(dt+toff+'+0000', '%Y%m%d%H%M%z')
        toff_jst = toff_utc.astimezone(datetime.timezone(datetime.timedelta(hours=+9)))
        toff_str = datetime.datetime.strftime(toff_jst,'%H%M')
        window.find_element('TIME_OFF').update(value=toff_str)


def adxImport(fn):
    global xmldata
    with open(fn, 'r', encoding='utf-8') as f:
        xmldata = ET.fromstring(f.read().encode('utf-8'))

# ADX出力
def adxExport(fn):
    global xmldata
    if not fn :
        fn = './'+datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9),'JST')).strftime('%Y%m%d%H%M%S')+'.adx'
    s = ET.tostring(xmldata, 'utf-8').decode('utf-8')
    s = re.sub('^\s+','',s,flags=re.MULTILINE)
    s = re.sub('\n','',s,flags=re.MULTILINE)
    doc = minidom.parseString(s)
    with open(fn, 'x', encoding='utf-8') as f:
        doc.writexml(f, encoding='utf-8', newl='\n', indent='', addindent='  ')

# ADX at Shift-JIS
def adxJisExport(fn):
    global xmldata
    if not fn :
        fn = './'+datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9),'JST')).strftime('%Y%m%d%H%M%S')+'.adx'
    s = ET.tostring(xmldata, 'utf-8').decode('utf-8')
    s = re.sub('^\s+','',s,flags=re.MULTILINE)
    s = re.sub('\n','',s,flags=re.MULTILINE)
    doc = minidom.parseString(s)
    with open(fn, 'x', encoding='shift-jis') as f:
        doc.writexml(f, encoding='shift-jis', newl='\n', indent='', addindent='  ')

# データ追加（末尾）
def data_add(window):
    global xmldata
    d = {}
    for k in inputList:
        v = window.find_element(k).get()
        if v : d[k] = v

    d = dataformat(d)
    if not d : return

    records = xmldata.find('./RECORDS')
    newrecord = ET.SubElement(records, 'RECORD')
    for k,v in d.items():
        elem = ET.SubElement(newrecord, k)
        elem.text = v

# データ更新
def data_update(window):
    global xmldata, show_number
    d = {}
    for k in inputList:
        v = window.find_element(k).get()
        d[k] = v

    d = dataformat(d)
    if not d : return

    records = xmldata.findall('./RECORDS/RECORD')
    record = records[show_number - 1]
    for e in [c for c in record]:
        record.remove(e)
    for k,v in d.items():
        if v:
            elem = ET.SubElement(record, k)
            elem.text = v

# データ追加、更新時のチェック、時刻修正
def dataformat(d):
    if not 'CALL' in d:
        SG.popup('C/Sが入力されていません')
        return
    if not 'QSO_DATE' in d:
        SG.popup('QSO Dateが入力されていません')
        return
    if not 'TIME_ON' in d:
        SG.popup('Time Onが入力されていません')
        return
    if not 'RST_RCVD' in d:
        SG.popup('RST Rcvdが入力されていません')
        return
    if not 'RST_SENT' in d:
        SG.popup('RST Sentが入力されていません')
        return

    ton_jst = datetime.datetime.strptime(d['QSO_DATE']+d['TIME_ON']+'+0900', '%Y%m%d%H%M%z')
    ton_utc = ton_jst.astimezone(datetime.timezone(datetime.timedelta(hours=0)))
    d['TIME_ON'] = datetime.datetime.strftime(ton_utc,'%H%M')

    if 'TIME_OFF' in d:
        toff_jst = datetime.datetime.strptime(d['QSO_DATE']+d['TIME_OFF']+'+0900', '%Y%m%d%H%M%z')
        toff_utc = toff_jst.astimezone(datetime.timezone(datetime.timedelta(hours=0)))
        d['TIME_OFF'] = datetime.datetime.strftime(toff_utc,'%H%M')

    d['QSO_DATE'] = datetime.datetime.strftime(ton_utc,'%Y%m%d')

    return d



if __name__ == '__main__':
    main()

