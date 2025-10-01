import os
import schedule
import time
def run_all_tasks():
    os.system("python scripts/task1.py")
    os.system("python scripts/task2.py")
    os.system("python scripts/task3.py")
    print("All tasks executed successfully.")
schedule.every().day.at("09:00").do(run_all_tasks)
while True:
    schedule.run_pending()
    time.sleep(60)