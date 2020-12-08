import requests
import lxml.html
import docx2txt
import os
import re
doc_folder_path = './'


AREA_CODE_LIST = []
AREA_CODES = {}

def download_doc():
    '''urlからファイルをダウンロードする
    '''
    # 総務省のページから市外局番一覧（WORD）をダウンロード
    response = requests.get('https://www.soumu.go.jp/main_sosiki/joho_tsusin/top/tel_number/shigai_list.html')
    response.encoding = 'shift-jis'
    html = lxml.html.fromstring(response.text.encode('shift-jis'))
    # HTML DOM
    a_elm =  html.xpath('//a[contains(text(), "WORD版はこちら")]')[0]
    word_link = 'https://www.soumu.go.jp' + a_elm.attrib['href']
    # ファイル名
    filename = word_link[word_link.rfind('/') + 1:]
    # wordデータ
    word_data = requests.get(word_link)
    # ローカルに保存する
    file_path = os.path.join(doc_folder_path, filename)
    with open(file_path, 'wb') as f:
        f.write(word_data.content)
    # 保存したファイルへのパスを返す
    return file_path


def get_doc_text(file):
    '''WORD(.doc)からテキストを取得する
    '''
    if file.endswith('.docx'):
       text = docx2txt.process(file)
       return text
    elif file.endswith('.doc'):
       # converting .doc to .docx
       doc_file = file
       docx_file = file + 'x'
       if not os.path.exists(docx_file):
          os.system('antiword ' + doc_file + ' > ' + docx_file)
          with open(docx_file) as f:
             text = f.read()
          os.remove(docx_file) #docx_file was just to read, so deleting
       else:
          # already a file with same name as doc exists having docx extension,
          # which means it is a different file, so we cant read it
          print('Info : file with same name of doc exists having docx extension, so we cant read it')
          text = ''
       return text


def parse_kukaku(text):
    '''番号区画をパースする
    '''
    pref = re.findall(r'\w{2,3}[都道府県]', text)[0]
    kukaku_items = text.split('、')
    items = []
    flg_kakko = False
    current_item = {}
    for item in kukaku_items:
        item = item.replace(pref, '')
        subitems_type = 'limit' if 'に限る' in item else 'ignore'
        if '（' in item and '）' in item:
            subitems = item.split('（')[1].replace('を除く。）','').replace('に限る。','').split('及び')
            items.append({'kukaku_name':item.split('（')[0], subitems_type:subitems})
        elif '（' in item:
            subitems = []
            current_item = {}
            current_item['kukaku_name'] = item.split('（')[0]
            subitems.append(item.split('（')[1])
            flg_kakko = True
        elif '）' in item:
            flg_kakko = False
            _items = item.replace('を除く。）','').replace('に限る。）','').split('及び')
            subitems.extend(_items)
            current_item[subitems_type] = subitems
            items.append(current_item)
        elif flg_kakko:
            subitems.append(item)
        else:
            items.append({'kukaku_name':item.split('（')[0]})
    return pref, items


def find_area_code_by_address(address):
    '''住所から市外局番情報を取得する
    '''
    area_codes = []
    # 番号区画が住所に含まれるものを取得
    for area_code in AREA_CODE_LIST:
        if area_code['番号区画'] in address:
            area_codes.append(area_code)
    # 除くに含まれる地名がある場合は除外
    for area_code in area_codes:
        for ignore in area_code['ignores']:
            if ignore in address:
                area_codes.remove(area_code)

    results = []
    # 限る。に含まれる地名がある場合は、ある場合のみ結果に追加
    for area_code in area_codes:
        if len(area_code['limits']) > 0:
            for limit in area_code['limits']:
                if limit in address:
                    results.append(area_code)
        else:
            results.append(area_code)
    return results


def find_area_code_by_phone_number(phone_number):
    '''電話番号から市外局番情報を取得する
    '''
    phone_num = re.sub(r"\D", "", phone_number)
    # 番号区画が住所に含まれるものを取得
    length = 5
    while True:
        if phone_num[:length] in AREA_CODES:
            return AREA_CODES[phone_num[:length]]
        length += -1
        if length < 2:
            return []


def get_area_codes(text):
    '''テキストからエリアデータを取得する
    '''
    result = []
    # 行
    lines = text.split('\n')
    current_data = {}
    datas = []
    for line in lines:
        # '|'で分割する。両端は不要
        items =[x.strip() for x in line.split('|')][1:-1]
        # 必要なデータが出てくるまでパス
        if len(items) < 2 or items[1] == '' or items[0] == '番号':
            continue

        # 番号区画を読み取る
        if items[0] == '':
            current_data['番号区画'] += items[1]
        else:
            # 区切り（現在のデータを確定する）
            if len(current_data.keys()) > 0:
                # 番号区画を正規化する
                pref, kukakus = parse_kukaku(current_data['番号区画'])
                current_data['都道府県'] = pref
                current_data['番号区画_json'] = kukakus
                datas.append(current_data)
            # 新規データを作成
            current_data = {}
            current_data['番号区画コード'] = items[0]
            current_data['番号区画'] = items[1]
            current_data['市外局番'] = items[2]
            current_data['市内局番'] = items[3]

    # 正規化
    for data in datas:
        for kukaku in data['番号区画_json']:
            res = {
                '市外局番': '0' + data['市外局番'],
                '番号区画コード': data['番号区画コード'],
                '番号区画': kukaku['kukaku_name'],
                'ignores': kukaku.get('ignore', []),
                'limits': kukaku.get('limit', []),
                '市内局番': data['市内局番'],
                '都道府県': data['都道府県'],
                '番号区画raw': data['番号区画']
            }
            result.append(res)
    return result


def load_area_code():
    '''市外局番情報をを読み込む
    '''
    global AREA_CODE_LIST, AREA_CODES
    # ダウンロードする
    file_path = download_doc()
    # docからテキストを取得する
    doc_text = get_doc_text(file_path)
    AREA_CODE_LIST = get_area_codes(doc_text)
    for area_code in AREA_CODE_LIST:
        if area_code['市外局番'] not in  AREA_CODES:
            AREA_CODES[area_code['市外局番']] = []
        AREA_CODES[area_code['市外局番']].append(area_code)


if __name__ == "__main__":
    load_area_code()
    import time
    st = time.time()
    res = find_area_code_by_address('鹿児島県鹿児島市')
    print(time.time()-st)
    print(res)
    st = time.time()
    res = find_area_code_by_phone_number('099-234-5678')
    print(time.time()-st)
    print(res)