from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import base64
import pprint 
import re
import csv
from datetime import datetime as dt
import locale

locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')


def find_text_start_from(keyword,text):
   search = keyword +".+"
   result = re.search(search, text)
   if result == None:
       return None
   else:
       return result.group(0).replace(keyword,"").strip('\r')


def decode_base64url_data(data):
    """
    base64url のデコード
    """
    decoded_bytes = base64.urlsafe_b64decode(data)
    decoded_message = decoded_bytes.decode("UTF-8")
    return decoded_message


def list_difference(list1, list2):
    result = list1.copy()
    for value in list2:
        if value in result:
            result.remove(value)

    return result

class GmailAPI:
    """
    GmailAPIに接続
    """
    def __init__(self):
        # If modifying these scopes, delete the file token.json.
        self._SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'


    def ConnectGmail(self):
        store = file.Storage('token.json')
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('credentials.json', self._SCOPES)
            creds = tools.run_flow(flow, store)
        service = build('gmail', 'v1', http=creds.authorize(Http()))

        return service


    def GetMessageList(self,DateFrom,DateTo,MessageFrom):

        #APIに接続
        service = self.ConnectGmail()

        MessageList = []

        query = ''
        # 検索用クエリを指定する
        if DateFrom != None and DateFrom !="":
            query += 'after:' + DateFrom + ' '
        if DateTo != None  and DateTo !="":
            query += 'before:' + DateTo + ' '
        if MessageFrom != None and MessageFrom !="":
            query += 'From:' + MessageFrom + ' '

        # メールIDの一覧を取得する(最大1000件)
        messageIDlist = service.users().messages().list(userId='me',maxResults=10000,q=query).execute()
        #該当するメールが存在しない場合は、処理中断
        if messageIDlist['resultSizeEstimate'] == 0: 
            print("Message is not found")
            return MessageList
        #メッセージIDを元に、メールの詳細情報を取得
        for message in messageIDlist['messages']:
            row = {}
            row['ID'] = message['id']
            MessageDetail = service.users().messages().get(userId='me',id=message['id']).execute()
       
            for header in MessageDetail['payload']['headers']:
                #日付、送信元、件名を取得する
                if header['name'] == 'Date':
                    row['Date'] = header['value'] 
                    #print(row['Date'])
                elif header['name'] == 'From':
                    row['From'] = header['value']
                elif header['name'] == 'Subject':
                    row['Subject'] = header['value']

            # bodyじゃなくてpartsに入っている場合もある
            if MessageDetail['payload']['body']['size'] != 0:
                decoded = decode_base64url_data(MessageDetail['payload']['body']['data'])
            else:
                decoded = decode_base64url_data(MessageDetail['payload']['parts'][0]['body']['data'])

            last_name = find_text_start_from("■名前（姓）：",decoded)
            first_name = find_text_start_from("■名前（名）：",decoded)
            name = str(last_name) + str(first_name)
            email = find_text_start_from("■メールアドレス：",decoded)
            phone = find_text_start_from("■電話番号：",decoded)
            num = find_text_start_from("■予約番号：",decoded)
            date = str(find_text_start_from("■利用日時：",decoded))
            count = find_text_start_from("■予約数：",decoded)

            
            if date != 'None':
                fix_date = date.split('～')[0]
                # 2021/02/27(土) 09:30～10:00
                datetime_date = dt.strptime(fix_date,'%Y/%m/%d(%a) %H:%M')
                # print(datetime_date)

                data = {'name': name,
                        'email': email,
                        'phone': phone,
                        'num': num,
                        'date': date,
                        'datetime_date': datetime_date,
                        'count': count

                        }

                MessageList.append(data)
            

        return MessageList


if __name__ == '__main__':
    test = GmailAPI()
    #パラメータは、任意の値を指定する
    messages = test.GetMessageList(DateFrom='2018-01-01',DateTo='2022-02-10',MessageFrom='reservation@airreserve.net')
    cancels = test.GetMessageList(DateFrom='2018-01-01',DateTo='2022-02-10',MessageFrom='reservation_cancel@airreserve.net')

    result = list_difference(messages, cancels)



    #結果を出力
    # ラベル作成用テストデータ
    save_dict = {'name': '南宏樹', 'email': 'mkoki0610@gmail.com', 'phone': '09019788665', 'num': '11ZEZ6QXJ', 'date': '2020', 'datetime_date': '2020', 'count': '2'}

    save_row = {}

    with open('test.csv','w') as f:
        writer = csv.DictWriter(f, fieldnames=save_dict.keys(),delimiter=",",quotechar='"')
        writer.writeheader()

        k1 = list(save_dict.keys())[0]
        length = len(save_dict[k1])

        for message in result:
            writer.writerow(message)





 

        