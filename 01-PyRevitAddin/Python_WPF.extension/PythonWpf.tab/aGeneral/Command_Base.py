# -*- coding: utf-8 -*-
# DelegateCommand.py
# lớp DelegateCommand để thực hiện các lệnh trong WPF
# có tác dụng kết nối các hàm trong MainViewModel với các sự kiện trong WPF
# có thể thực hiện các lệnh và kiểm tra điều kiện thực hiện lệnh


from System.Windows.Input import ICommand
from System import EventHandler, EventArgs

class DelegateCommand(ICommand):
    def __init__(self, execute, can_execute=None):
        self.execute = execute
        self.can_execute = can_execute
        self._canExecuteChangedHandlers = []

    def Execute(self, parameter):
        if self.execute:
            self.execute(parameter)  # Truyền parameter vào hàm execute

    def CanExecute(self, parameter):
        if callable(self.can_execute):
            return self.can_execute(parameter)
        return True  # Mặc định lệnh luôn thực thi được nếu không có can_execute

    def add_CanExecuteChanged(self, handler):
        self._canExecuteChangedHandlers.append(handler)

    def remove_CanExecuteChanged(self, handler):
        if handler in self._canExecuteChangedHandlers:
            self._canExecuteChangedHandlers.remove(handler)

    def RaiseCanExecuteChanged(self):
        for handler in self._canExecuteChangedHandlers:
            handler(self, EventArgs.Empty)