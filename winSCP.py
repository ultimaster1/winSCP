import selenium.webdriver as webdriver
import logging
import time
import paramiko
import keyboard
from selenium.webdriver.common.by import By
import os
from tkinter import *
from tkinter import messagebox as mb
from tkinter import filedialog, simpledialog
import win32gui,win32con
from pykeepass import PyKeePass
import pathlib



class automotation_server_conn():
    def __init__(self):
        self.driver = None
        self.main_window = None
        self.password = None
        self.login = None
        self.kp = None
        self.path_maneger = None
        self.content = {}
        self.hwds = []


    def winEnumHandler(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            self.hwds.append(hwnd)


    def clearBoard(self):
        win32gui.EnumWindows(self.winEnumHandler, None)
        for i in self.hwds:
            try:
                win32gui.ShowWindow(i, win32con.SW_MINIMIZE)
            except:
                pass
        self.hwds = []


    def file_len(self,fname):
        with open(fname) as f:
            for i, l in enumerate(f):
                pass
        return i + 1


    def ask_file(self, name):
        f = open('pathes.txt', "a")
        a = filedialog.askopenfilename()
        f.write(name + ': ' + a + "\n")


    def remote_PC(self):
        self.clearBoard()
        os.startfile(self.content['PowerShell_path'])
        time.sleep(5)
        keyboard.write(
            "ssh -i '" + str(self.content['KeyDesktopFile_path']) + "'-L 33389:192.168.100.1:3389 -l "
            + self.login + " 82.202.167.88")
        keyboard.press('enter')
        time.sleep(5)
        keyboard.write(self.kp.find_entries_by_title('keys_etc', first=True).password)
        time.sleep(3)
        keyboard.press('enter')
        time.sleep(2)
        self.clearBoard()
        try:
            self.driver.quit()
            self.driver = webdriver.Remote(
                command_executor='http://localhost:9999',
                desired_capabilities={
                    "app": self.content['mstsc_path'].as_posix()
                })  # Вводим путь к файл и порт на котором будет работать эта хрень
        except:
            self.driver = webdriver.Remote(
                command_executor='http://localhost:9999',
                desired_capabilities={
                    "app": self.content['mstsc_path'].as_posix()
                })  # Вводим путь к файл и порт на котором будет работать эта хрень
        time.sleep(5)
        keyboard.press('backspace')
        keyboard.write('localhost:33389')
        self.driver.find_element_by_name('Подключить').click()
        time.sleep(5)
        keyboard.write(self.login)
        time.sleep(3)
        keyboard.press('tab')
        time.sleep(3)
        keyboard.write(self.kp.find_entries_by_title('remote_pc', first=True).password)
        time.sleep(3)
        keyboard.press('enter')


    def winSCP(self):
        self.clearBoard()
        try:
            self.driver.quit()
            self.driver = webdriver.Remote(
                command_executor='http://localhost:9999',
                desired_capabilities={
                    "app": self.content['WinSCP_path'].as_posix()
                })  # Вводим путь к файл и порт на котором будет работать эта хрень
        except:
            self.driver = webdriver.Remote(
                command_executor='http://localhost:9999',
                desired_capabilities={
                    "app": self.content['WinSCP_path'].as_posix()
                })  # Вводим путь к файл и порт на котором будет работать эта хрень
        time.sleep(5)
        self.driver.find_elements(By.CLASS_NAME,"TEdit")[1].send_keys('82.202.167.88')
        self.driver.find_elements(By.CLASS_NAME,"TEdit")[0].send_keys(self.login)
        self.driver.find_element_by_name('Ещё…').click()
        self.driver.find_element_by_name('Туннель').click()
        self.driver.find_element_by_name('Соединяться через SSH-туннель').click()
        self.driver.find_elements(By.CLASS_NAME,"TEdit")[1].send_keys('82.202.167.88')
        self.driver.find_elements(By.CLASS_NAME,"TEdit")[0].send_keys(self.login)
        self.driver.find_element_by_name('...').click()
        self.driver.find_element_by_name('Имя файла:').click()
        time.sleep(4)
        keyboard.write(str(self.content['KeyDesktopPpk_path']))
        self.driver.find_element_by_name('Открыть').click()
        self.driver.find_element_by_name('Аутентификация').click()
        self.driver.find_element_by_name('...').click()
        self.driver.find_element_by_name('Имя файла:').click()
        time.sleep(4)
        keyboard.write(str(self.content['KeyPredtechPpk_path']))
        self.driver.find_element_by_name('Открыть').click()
        self.driver.find_element_by_name('ОК').click()


    def conn_to_server(self):
        self.clearBoard()
        logging.basicConfig(stream=sys.stderr, level=logging.DEBUG)
        host = '82.202.167.88'
        user = self.login
        secret = self.kp.find_entries_by_title('keys_etc', first=True).password
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(hostname=host, username=user, port=22, password=secret,
                    key_filename=self.content['KeyDesktopFile_path'].as_posix())
        vmtransport = ssh.get_transport()
        dest_addr = ('192.168.100.1', 22)  # edited#
        local_addr = ('82.202.167.88', 22)  # edited#
        vmchannel = vmtransport.open_channel("direct-tcpip", dest_addr, local_addr)

        jhost = paramiko.SSHClient()
        jhost.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        jhost.connect(hostname='192.168.100.1', username=user, password=secret, sock=vmchannel,
                      key_filename=self.content['KeyDesktopFile_path'].as_posix())

        while not keyboard.is_pressed('esc'):
            stdin, stdout, stderr = jhost.exec_command(input())
            print(stdout)
            time.sleep(0.1)


    def mySQL(self):
        self.clearBoard()
        os.startfile(self.content['PowerShell_path'])
        time.sleep(5)
        keyboard.write(r"ssh - i '" + self.content['KeyPredtechPpk_path'] + "'- L 3306: 192.168.100.1: 3306 - l" +
                       self.login + " 82.202.167.88")
        time.sleep(5)
        keyboard.press('enter')


    def new_window(self):
        self.main_window.destroy()
        if os.path.isfile("pathes.txt") and os.stat("pathes.txt").st_size != 0:
            for i in open('pathes.txt').readlines():
                if 'PowerShell_path' in i:
                    self.content['PowerShell_path'] = pathlib.PureWindowsPath(i.split(' : ')[1].replace('\n', ''))
                elif 'KeyDesktopFile_path'  in i:
                    self.content['KeyDesktopFile_path'] = pathlib.PureWindowsPath(i.split(' : ')[1].replace('\n', ''))
                elif 'KeyPredtechFile_path' in i:
                    self.content['KeyPredtechFile_path'] = pathlib.PureWindowsPath(i.split(' : ')[1].replace('\n', ''))
                elif 'KeyDesktopPpk_path' in i:
                    self.content['KeyDesktopPpk_path'] = pathlib.PureWindowsPath(i.split(' : ')[1].replace('\n', ''))
                elif 'KeyPredtechPpk_path' in i:
                    self.content['KeyPredtechPpk_path'] = pathlib.PureWindowsPath(i.split(' : ')[1].replace('\n', ''))
                elif 'WinSCP_path' in i:
                    self.content['WinSCP_path'] = pathlib.PureWindowsPath(i.split(' : ')[1].replace('\n', ''))
                elif 'mstsc_path' in i:
                    self.content['mstsc_path'] = pathlib.PureWindowsPath(i.split(' : ')[1].replace('\n', ''))
            self.main_window = Tk()
            self.main_window.geometry("600x500")
            label = Label(self.main_window, text="Выберите опцию")
            label.pack(pady=10)
            mod_0 = Button(self.main_window, command=self.remote_PC,
                           text="Удаленный рабочий стол")
            mod_1 = Button(self.main_window, command=self.mySQL,
                           text="MySQL")
            mod_2 = Button(self.main_window, command=self.winSCP,
                           text="Файловая система")
            mod_3 = Button(self.main_window, command=self.conn_to_server,
                           text="Командная строка")

            mod_0.pack(pady=10)
            mod_1.pack(pady=10)
            mod_2.pack(pady=10)
            mod_3.pack(pady=10)
            self.main_window.mainloop()
        else:
            if os.path.isfile("pathes.txt"):
                file = open('pathes.txt', 'w+')
            self.main_window = Tk()
            self.main_window.geometry("600x500")
            label = Label(self.main_window, text="Заполните пути")
            label.pack(pady=10)
            path_1 = Button(self.main_window, command=lambda: self.ask_file('PowerShell_path'),text="Заполниить путь к ярлыку PowerShell")
            path_2 = Button(self.main_window, command=lambda: self.ask_file("KeyDesktopFile_path"), text="Заполниить путь к ключу .file desktop")
            path_3 = Button(self.main_window, command=lambda: self.ask_file("KeyPredtechFile_path"), text="Заполниить путь к ключу .file predtech")
            path_4 = Button(self.main_window, command=lambda: self.ask_file("KeyDesktopPpk_path"), text="Заполниить путь к ключу .ppk desktop")
            path_5 = Button(self.main_window, command=lambda: self.ask_file("KeyPredtechPpk_path"),text="Заполниить путь к ключу .ppk predtech")
            path_6 = Button(self.main_window, command=lambda: self.ask_file("WinSCP_path"), text="Заполниить путь к exe WinSCP")
            path_7 = Button(self.main_window, command=lambda: self.ask_file("mstsc_path"), text="Заполниить путь к exe удаленному рабочему столу")
            activate = Button(self.main_window, command=self.new_window, text="Запуск")

            path_1.pack()
            path_2.pack()
            path_3.pack()
            path_4.pack()
            path_5.pack()
            path_6.pack()
            path_7.pack()
            activate.pack()


    def pass_manager(self):
        self.path_maneger = filedialog.askopenfilename()
        if self.password and self.login:
            self.kp = PyKeePass(self.path_maneger, password=self.password)
            self.new_window()


    def Login(self):
        self.login = simpledialog.askstring('Login', 'Enter login')
        if self.path_maneger and self.password:
            self.kp = PyKeePass(self.path_maneger, password=self.password)
            self.new_window()


    def main_pass(self):
        self.password = simpledialog.askstring('Password','Enter password',show='*')
        if self.path_maneger and self.login:
            self.kp = PyKeePass(self.path_maneger, password=self.password)
            self.new_window()


if __name__ == '__main__':
    AServer = automotation_server_conn()
    AServer.main_window = Tk()
    AServer.main_window.geometry("350x500")
    label = Label(AServer.main_window, text="Подключитесь к база данных с паролями")
    label.pack(pady=10)
    mod_0 = Button(AServer.main_window,command = AServer.pass_manager, text="Путь к базе данных паролей")
    mod_1 = Button(AServer.main_window, command=AServer.Login, text="Введите логин")
    mod_2 = Button(AServer.main_window,command = AServer.main_pass, text="Введите пароль")
    mod_0.pack(pady=10)
    mod_1.pack(pady=10)
    mod_2.pack(pady=10)
    AServer.main_window.mainloop()