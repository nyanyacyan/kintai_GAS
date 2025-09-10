import configparser
import logging
import subprocess
import sys
import time
from datetime import datetime

import pytz
import schedule

# ログ設定
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# ファイルハンドラ
file_handler = logging.FileHandler('scheduled_launcher.log', encoding='utf-8')
file_handler.setLevel(logging.INFO)

# コンソールハンドラ
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# フォーマット
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# ハンドラをロガーに追加
logger.addHandler(file_handler)
logger.addHandler(console_handler)


def load_config(config_file: str = 'config.ini') -> tuple[list[str], str, str, list[str]]:
    config = configparser.ConfigParser()
    with open(config_file, 'r', encoding='utf-8') as f:  # エンコーディングを指定
        config.read_file(f)
    times = config["Schedule"]["time"].split(',')
    days = config["Schedule"]["days"].split(',')
    return times, config["Schedule"]["program_path"], config['Schedule']['timezone'], days


def reformat_shcedule_time(schedule_time: str, timezone: str) -> str:
    tz = pytz.timezone(timezone)
    now = datetime.now(tz)
    _schedule_time = datetime.strptime(schedule_time, "%H:%M").time()
    schedule_datetime = datetime.combine(now.date(), _schedule_time)
    return schedule_datetime.strftime("%H:%M")


def run_program(program_path: str) -> None:
    try:
        python_executable = sys.executable
        subprocess.run([python_executable, program_path], check=True)
        logging.info(f"プログラムを実行しました: {sys.argv[0]} -> {program_path}")
    except subprocess.CalledProcessError as e:
        logging.error(f"プログラムの実行中にエラーが発生しました: {e}")


if __name__ == "__main__":
    schedule_times, program_path, timezone, days_of_week = load_config()

    for day, time_str in zip(days_of_week, schedule_times):
        reformat_schedule_datetime = reformat_shcedule_time(time_str, timezone)
        logging.info(f"{reformat_schedule_datetime} にプログラム: {program_path} を起動するようにスケジュールされました。")
        if day.lower() == 'every':
            schedule.every().day.at(reformat_schedule_datetime).do(run_program, program_path)
        elif day.lower() == 'monday':
            schedule.every().monday.at(reformat_schedule_datetime).do(run_program, program_path)
        elif day.lower() == 'tuesday':
            schedule.every().tuesday.at(reformat_schedule_datetime).do(run_program, program_path)
        elif day.lower() == 'wednesday':
            schedule.every().wednesday.at(reformat_schedule_datetime).do(run_program, program_path)
        elif day.lower() == 'thursday':
            schedule.every().thursday.at(reformat_schedule_datetime).do(run_program, program_path)
        elif day.lower() == 'friday':
            schedule.every().friday.at(reformat_schedule_datetime).do(run_program, program_path)
        elif day.lower() == 'saturday':
            schedule.every().saturday.at(reformat_schedule_datetime).do(run_program, program_path)
        elif day.lower() == 'sunday':
            schedule.every().sunday.at(reformat_schedule_datetime).do(run_program, program_path)
        else:
            logging.error(f"無効な曜日: {day}")

    while True:
        schedule.run_pending()
        time.sleep(1)
