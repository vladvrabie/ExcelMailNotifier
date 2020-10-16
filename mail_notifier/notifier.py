import argparse
import os
import pathlib
import smtplib
import ssl
import sys
import time
from datetime import datetime

try:
    import keyring

    from mail_notifier.excel_reader import read_excel
    from mail_notifier import logger
except ImportError as ie:
    try:
        home = str(pathlib.Path.home())
        log_path = os.path.join(home, "log_file.log")
        with open(log_path, 'a') as import_logger:
            import_logger.write(f'[{str(datetime.today())}]: {str(ie)}\n')
    except Exception:
        pass
    finally:
        sys.exit(0)

if sys.version_info < (3, 9):
    from typing import List
else:
    List = list


def send_email(list_to_send: List[str]) -> None:
    host = 'smtp.gmail.com'
    ssl_google_port = 465
    context = ssl.create_default_context()

    service = '...'
    sender_email = '...'
    receiver_email1 = '...'
    receiver_email2 = '...'
    receivers = [receiver_email1, receiver_email2]

    with smtplib.SMTP_SSL(
        host,
        ssl_google_port,
        context=context
    ) as server:
        if keyring.get_password(service, sender_email) is not None:
            server.login(sender_email,
                         keyring.get_password(service, sender_email))

            message_body = '\n'.join(list_to_send)

            message = (
                f'From: {sender_email}\n'
                f'To: {receiver_email1},{receiver_email2}\n'
                'Subject: Notificare expirare acte\n\n\n'
                'Urmatoarele randuri din tabel au '
                'data de expirare in viitorul apropiat:\n\n'
                f'{message_body}'
                '\n\n\n  O zi frumoasa!'
            )

            logger.log(message.replace('\n', ' '))
            server.sendmail(sender_email, receivers, message)
            logger.log('Email sent')
        else:
            logger.log("Password was not received from keyring.",
                       logger.ERROR)


def try_send_daily(details_to_send: List[str]) -> bool:
    if len(details_to_send) < 2:
        # assuming it will contain the header row
        return True  # nothing to send

    times_retried = 0
    sleep_time = 5  # in seconds
    # no. of seconds in a day divided by sleep time
    max_tries = 24 * 60 * 60 // sleep_time
    while True:
        try:
            if times_retried >= max_tries:
                # checking if a day has passed
                # if computer is left to run overnight,
                # new details may be needed to be sent
                return False
            send_email(details_to_send)
            return True  # message sent
        except Exception as e:
            # message not sent
            logger.log(str(e), logger.ERROR)
            times_retried += 1
            time.sleep(sleep_time)


if __name__ == '__main__':
    try:
        time.sleep(2)  # to avoid weird throttling restart delay from nssm

        parser = argparse.ArgumentParser()
        parser.add_argument(
            'excel_path',
            help='Path to the Excel file to monitor.'
        )
        parser.add_argument(
            'columns_to_monitor',
            nargs='+',
            type=int,
            help='Columns of dates to monitor.'
        )
        parser.add_argument(
            '-d',
            '--date-format',
            nargs='*',
            default='%d.%m.%Y',
            help='Format of dates in column to monitor. Default: %d.%m.%Y'
        )
        parser.add_argument(
            '-t',
            '--table-head',
            type=int,
            default=0,
            help='Row on which the table head can be found.'
        )

        args = parser.parse_args()
        excel_path = args.excel_path
        cols_to_monitor = args.columns_to_monitor
        date_format = args.date_format
        table_head_row = args.table_head

        details_to_send = read_excel(excel_path, cols_to_monitor,
                                     table_head_row, date_format)
        print(details_to_send)
        message_sent = try_send_daily(details_to_send)
        print(message_sent)

        while not message_sent:
            # a day passed, so we need to re-read the excell
            details_to_send = read_excel(excel_path, cols_to_monitor,
                                         table_head_row, date_format)
            message_sent = try_send_daily(details_to_send)

    except Exception as e:
        try:
            logger.log(str(e), logger.ERROR)
        except Exception:
            pass  # failed on changing current directory to mail_notifier
            # or couldn't import logger
