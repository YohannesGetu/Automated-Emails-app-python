import yagmail
import pandas
from datetime import datetime, timedelta
import time
from news import NewsFeed


def send_email(row, today, yesterday):
    news_feed = NewsFeed(interest=row['interest'], from_date=yesterday, to_date=today)
    email = yagmail.SMTP(user="pythonprocourse1@gmail.com", password="python_pro_course_1")
    email.send(to=row['email'],
               subject=f"Your {row['interest']} news for today!",
               contents=f"Hi {row['name']}\n See what's on about {row['interest']} today. \n{news_feed.get()} "
                        f"\nYohannes G")


while True:
    if datetime.now().hour == 15 and datetime.now().minute == 20:
        df = pandas.read_excel('people.xlsx')

        yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        today = datetime.now().strftime('%Y-%m-%d')

        for index, row in df.iterrows():
            send_email(row, today, yesterday)
    time.sleep(60)
