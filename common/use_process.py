import time
import multiprocessing
import keyboard


# 进程任务的示例
def process_task():
    while True:
        print("进程正在运行...")
        time.sleep(1)


# 进程管理类
class ProcessManager:
    def __init__(self):
        self.process = None
        self.is_running = False
        self.is_paused = False
        self.pause_event = multiprocessing.Event()

    # 启动进程
    def start(self):
        if not self.is_running:
            self.process = multiprocessing.Process(target=self.run_task)
            self.process.start()
            self.is_running = True
            print("进程启动成功")
        else:
            print("进程已经在运行")

    # 运行任务
    def run_task(self):
        while True:
            if self.is_paused:
                self.pause_event.wait()  # 等待直到暂停被解除
            process_task()

    # 暂停进程
    def pause(self):
        if self.is_running and not self.is_paused:
            self.is_paused = True
            print("进程已暂停")
        else:
            print("进程未启动或已经暂停")

    # 恢复进程
    def resume(self):
        if self.is_paused:
            self.is_paused = False
            self.pause_event.set()  # 解除暂停
            print("进程已恢复")
        else:
            print("进程没有暂停")

    # 结束进程
    def stop(self):
        if self.is_running:
            self.process.terminate()
            self.process.join()  # 确保进程结束
            self.is_running = False
            self.is_paused = False
            print("进程已结束")
        else:
            print("进程未启动")


# 使用进程
def operation_manage():
    manager = ProcessManager()
    manager.start()
    while True:
        if keyboard.is_pressed('esc'):
            print('Ctrl+C被按下')
            manager.pause()
            manager.stop()
            break


# if __name__ == "__main__":
#     manager = ProcessManager()
#
#     # 启动进程
#     manager.start()
#     time.sleep(5)  # 运行 5 秒
#
#     # 暂停进程
#     manager.pause()
#     time.sleep(5)  # 暂停 5 秒
#
#     # 恢复进程
#     manager.resume()
#     time.sleep(5)  # 再运行 5 秒
#
#     # 停止进程
#     manager.stop()
