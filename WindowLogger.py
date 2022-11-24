#
#   Author: Jagjot Singh
#   Date:31/12/18 6:30 PM
#   Last Edit:30/01/19 4:02 AM
#
import time
from datetime import datetime
from platform import system
from win32api import GetComputerName, GetUserName
from win32gui import GetForegroundWindow, GetWindowText
from win32process import GetWindowThreadProcessId
from psutil import Process, pid_exists
from sys import argv

# This number should be handled with CARE!
# Increasing too much will make logger skip new windows
# but reduces CPU usage and vice-versa.
# Default set to 0.2 secs
REFRESH_RATE = 0.2

# This number used to set main loop's sleep rate.
# This number should not be changed.
LOOP_SLEEP_RATE = 0.001


def file_name_creator():
    """
    # Returns the name of the '.log' file in the format "user_name@computer_name_date.log"
    # GetUserName() is a function that returns 'user_name' as string
    # GetComputerName() is a function that returns 'computer_name' as string
    # datetime.today().date() is a function that returns current date as YYYY-MM-DD
    """
    return "{}@{}_{}.log".format(GetUserName(), GetComputerName(), datetime.today().date())


def logger(file_name):
    """
    Main logger function that add entries to the '.log' file.
    """
    while True:
        # To reduce the number of iterations when 'pid' doesn't exist!
        time.sleep(LOOP_SLEEP_RATE)

        current_window = GetForegroundWindow()
        pid = GetWindowThreadProcessId(current_window)[1]

        if pid_exists(pid):

            # To create new file when logger() is running but date is changed i.e next day.
            if file_name[-14:-4] != datetime.today().date().__str__():
                file_name = file_name_creator()

            # Stores the title bar text in bytes format in order to store non ASCII characters.
            window_text = GetWindowText(current_window).encode("utf-8")

            # Stores information about a process in a dictionary format.
            process_info_list = Process(pid).as_dict()

            start_time = datetime.now()

            # Open log file to add first half of log entry.
            file = open(file_name, 'a')

            # Adding first half of log entry.
            file.write("\n" + "{:<8} | {:<30} | {:<11}".format(process_info_list['pid'], process_info_list['name'],
                                                               start_time.strftime("%I:%M:%S %p")))
            # Closing the log file.
            file.close()

            # print("\n" + "{:<8} | {:<30} | {:<11}".format(process_info_list['pid'], process_info_list['name'],
            #                                               start_time.strftime("%I:%M:%S %p")), end='')

            # Logic for not adding a new entry in log when user is on the same window/tab.
            next_window_text = GetWindowText(GetForegroundWindow()).encode("utf-8")
            while next_window_text == window_text:
                time.sleep(REFRESH_RATE)
                next_window_text = GetWindowText(GetForegroundWindow()).encode("utf-8")

            end_time = datetime.now()

            active_time = end_time - start_time

            # Open log file to add second half of log entry.
            file = open(file_name, 'a')

            # Adding second half of log entry.
            file.write(" | {:<11} | {:<9} | {} | {}".format(end_time.strftime("%I:%M:%S %p"), active_time.__str__(),
                                                            process_info_list['exe'], window_text))

            # Closing the log file.
            file.close()

            # print(" | {:<11} | {:<15} | {} | {}".format(end_time.strftime("%I:%M:%S %p"), active_time.__str__(),
            #                                             process_info_list['exe'], window_text, end=''))


# To make sure this script will run on 'Microsoft Windows OS'
if __name__ == '__main__' and system() in ['Windows', 'win32', 'win64']:
    if len(argv) > 1:
        txt = "This program is not meant to be run with any commandline arguments."
        print(txt + " \"readme.txt\" created.")
        file = open("readme.txt", 'w')
        file.write(
             txt + "\nGithub: https://github.com/singh-jagjot/WindowLogger"
        )
        file.close()
    else:
        logger(file_name_creator())
else:
    print("This program can only work on Microsoft Windows OS!")