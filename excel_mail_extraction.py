import pandas as pd


# item = "2022 수능특강 동영쌤 손필기 유형편"
item = "2020년 3월 모의고사 동영쌤 손필기 [고1, 고2]"
prefix = "/Users/dongho/Documents/Dbooks/배송엑셀/2020-고1,2-3월모의엑셀/발송처리2021"
suffix = ".xlsx"
# date_list = ['0323', '0324', '0325', '0326', '0329', '0330', '0331']
# date_list = ['0401', '0402']
# date_list = ['0405']
# date_list = ['0406']
# date_list = ['0407']
date_list = ['0408', '0409']


def excel_read(prefix, suffix):
    xlsx = pd.read_excel(prefix + suffix, sheet_name='발송처리', engine='openpyxl')
    return xlsx


def extraction_ids(xlsx, item):
    xl_ids = xlsx.loc[:, ['구매자ID']]
    xl_item = xlsx.loc[:, ['상품명']]
    ids = xl_ids.values
    items = xl_item.values

    ret_ids = []

    for i in range(len(ids)):
        if items[i] == item:
            ret_ids.append(ids[i])

    return ret_ids


def id_converter(mails, ids):
    for i in range(len(ids)):
        mails.append(ids[i][0]+"@naver.com")


if __name__ == "__main__":

    mails = []
    # 날짜별로 아이디 추출해서 한 리스트에 넣기
    for date in date_list:
        xlsx = excel_read(prefix, date+suffix)
        ids = extraction_ids(xlsx, item)
        id_converter(mails, ids)

    mails = set(mails)
    mails = list(mails)

    # 리스트에 있는 아이디를 100개 단위로 출력하기
    count = 0
    for mail in mails:
        print(mail, end=', ')
        count += 1
        if count == 100:
            print()
            count = 0

    print()
    print(f"총 {len(mails)} 개의 아이디를 메일로 바꿨습니다.")



