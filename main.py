import xlrd
import pyautogui as pa
import time
from datetime import datetime

time.sleep(2)

starting_index=0
url_x=518
url_y=60
done_x=508
done_y=502
time_lag=2

# print(pa.position())

def get_order_ids():
    loc = ("orders.xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    nr=sheet.nrows

    arr=[]
    for i in range(1,nr):
        try:
            date=sheet.cell_value(i, 2)[:10]
            date_object = datetime.strptime(date, '%Y-%m-%d')
            delta=datetime.today()-date_object
            diff=delta.days

            if(diff>8 and diff<40):
                arr.append(sheet.cell_value(i, 0))

        except Exception as e:
            print(str(e))
    return arr


def request_review(order_id):
    try:
        url="https://sellercentral.amazon.in/messaging/reviews?orderId="+str(order_id)+"&marketplaceId=A21TJRUUN4KGV"
        pa.click(url_x,url_y)
        pa.write(url)
        pa.press("enter")

        time.sleep(2*time_lag)
        pa.click(done_x,done_y)
    except Exception as e:
        print(str(e))



def main():
    li = get_order_ids()
    print("Number of Order Ids: "+str(len(li)))
    cnt=0
    for id in range(starting_index,len(li)):
        time.sleep(time_lag)
        request_review(li[id])
        print(id)
        cnt+=1
    print(str(cnt)+" Reviews Requested")

main()