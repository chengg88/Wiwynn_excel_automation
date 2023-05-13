from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.interval import IntervalTrigger
import time
from batch import Auto_excel

Auto_excel = Auto_excel()

def task():
    localtime = time.asctime(time.localtime(time.time()))
    Auto_excel.generate_report()
    print(localtime,": 生成季報表...")

scheduler = BlockingScheduler(timezone="Asia/Taipei")
trigger = IntervalTrigger(minutes=5, timezone="Asia/Taipei")
scheduler.add_job(task, trigger)

try:
    scheduler.start()
except (KeyboardInterrupt, SystemExit):
    scheduler.shutdown()

