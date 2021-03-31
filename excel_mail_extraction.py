import pandas as pd

prefix = "/Users/dongho/Documents/Dbooks/배송엑셀/2020-고1,2-3월모의엑셀/발송처리2021"
surfix = ".xlsx"
date_list = ['0323', '0324', '0325', '0326', '0329', '0330', '0331']


def excel_read(prefix, surfix):
    xlsx = pd.read_excel(prefix + surfix, sheet_name='발송처리', engine='openpyxl')
    return xlsx


def extraction_ids(xlsx):
    xl_ids = xlsx.loc[:, ['구매자ID']]
    ids = xl_ids.values
    return ids


def id_converter(mails, ids):
    for i in range(len(ids)):
        mails.append(ids[i][0]+"@naver.com")


if __name__ == "__main__":

    mails = []
    # 날짜별로 아이디 추출해서 한 리스트에 넣기
    for date in date_list:
        xlsx = excel_read(prefix, date+surfix)
        ids = extraction_ids(xlsx)
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



