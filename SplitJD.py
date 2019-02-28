import pandas as pd

def Split(text):
    # print(text)

    # print("=====================================================")

    # Responsibility
    ridx = text.find("Responsibilit")
    if ridx == -1:
        ridx = text.find("responsibilit")

    if ridx == -1:
        print("No Responsibilities found")

    # Qualification
    qidx = 2147483647
    if text.find("qualification") > 0:
        qidx = min(qidx, text.find("qualification"))
    if text.find("Qualification") > 0:
        qidx = min(qidx, text.find("Qualification"))
    elif text.find("require") > 0:
        qidx = min(qidx, text.find("require"))
    elif text.find("Require") > 0:
        qidx = min(qidx, text.find("Require"))
    if qidx == 2147483647:
        qidx = -1

    if qidx == -1:
        if (ridx == -1):
            return "", ""
        else:
            resp = text[ridx:]
    else:
        tmp = text.rfind("\n", 0, qidx)
        if tmp == -1:
            print("No new Line found!")
            resp = text[ridx:]
        else:
            resp = text[ridx:tmp]

    # print(resp)

    # print("=====================================================")

    tmpqidx = max(text.rfind("qualification"), text.rfind("Qualification"))
    # print(qidx)
    # print(tmpqidx)
    rqidx = max(tmpqidx, qidx)
    rblank = text.find("\n\n", rqidx)
    # print(rblank)
    if rblank - rqidx < 300:
        quali = text[qidx:]
    else:
        quali = text[qidx:rblank]
    # print(quali)

    return resp, quali

if __name__ == '__main__':
    df = pd.read_excel('output.xlsx')
    text = df['Job Description'][2]
    Split(text)