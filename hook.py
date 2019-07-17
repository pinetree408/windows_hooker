import mouse
import keyboard
import win32gui
import win32ui
import win32con
import win32api
import sys
import time
import os
import threading


class Hooker(object):
    def __init__(self, result_dir):
        self.mouse_events = []
        self.keyboard_events = []
        self.filename = None
        self.result_dir = result_dir
        if not os.path.exists(self.result_dir):
            os.makedirs(self.result_dir)

    def hook(self):
        mouse.hook(self.mouse_events.append)

        def _keyboard_hook(event):
            self.keyboard_events.append(event)
        keyboard.hook(_keyboard_hook)

    def unhook(self):
        mouse.unhook(self.mouse_events.append)
        keyboard.unhook_all()

    def get_event_from_raw(self, event, category):
        now = time.localtime(event.time)
        formatted_now = \
            str(now.tm_year) + "-" + str(now.tm_mon) + "-" +\
            str(now.tm_mday) + " " + str(now.tm_hour) + ":" +\
            str(now.tm_min) + ":" + str(now.tm_sec)
        obj = {
            "time": formatted_now,
            "process": win32gui.GetWindowText(win32gui.GetForegroundWindow()),
            "category": category,
            "type": "",
            "info": "",
        }

        if category == "keyboard":
            obj["type"] = event.event_type
            obj["info"] = event.name
        elif category == "mouse":
            class_name = event.__class__.__name__
            event_type = ""
            event_info = ""
            if class_name == "MoveEvent":
                event_type = "move"
                event_info = str(event.x)+","+str(event.y)
            elif class_name == "WheelEvent":
                event_type = "wheel"
                event_info = event.delta
            elif class_name == "ButtonEvent":
                event_type = "button"
                event_info = event.button+"-"+event.event_type
            obj["type"] = event_type
            obj["info"] = event_info

        return obj

    def write_log(self, obj):
        if self.filename is not None:
            with open(
                    self.result_dir + "/" + self.filename + ".txt",
                    "a"
                    ) as f:
                log_str =\
                    obj['time'] + "\t" + obj["process"] + "\t" +\
                    obj["category"] + "\t" + obj["type"] + "\t" +\
                    obj["info"] + "\n"
                f.write(log_str)

    def get_screen_capture(self):
        while True:
            cur_time = time.strftime(
                    "%Y%m%d_%H%M%S", time.localtime(time.time())
                    )

            hdesktop = win32gui.GetDesktopWindow()

            width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
            height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
            left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
            top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)

            desktop_dc = win32gui.GetWindowDC(hdesktop)
            img_dc = win32ui.CreateDCFromHandle(desktop_dc)

            mem_dc = img_dc.CreateCompatibleDC()

            screenshot = win32ui.CreateBitmap()
            screenshot.CreateCompatibleBitmap(img_dc, width, height)
            mem_dc.SelectObject(screenshot)

            mem_dc.BitBlt(
                    (0, 0), (width, height),
                    img_dc, (left, top), win32con.SRCCOPY
                    )

            screenshot.SaveBitmapFile(
                    mem_dc, self.result_dir + "/" + cur_time + ".png"
                    )

            mem_dc.DeleteDC()
            win32gui.DeleteObject(screenshot.GetHandle())

            time.sleep(1)

    def logging(self, filename):
        self.filename = filename
        screen_capture_thread = threading.Thread(
                target=self.get_screen_capture,
                args=()
                )
        screen_capture_thread.daemon = True
        screen_capture_thread.start()

        with open(self.result_dir + "/" + self.filename + ".txt", "a") as f:
            f.write("time\tprocess\tcategory\ttype\tinfo\n")
        while True:
            mouse._listener.queue.join()
            keyboard._listener.queue.join()

            logs = []
            for event in self.mouse_events:
                logs.append(self.get_event_from_raw(event, "mouse"))
            for event in self.keyboard_events:
                logs.append(self.get_event_from_raw(event, "keyboard"))

            for log in logs:
                print(log)
                self.write_log(log)
                sys.stdout.flush()

            self.mouse_events.clear()
            self.keyboard_events.clear()


if __name__ == "__main__":
    hooker = Hooker("results")
    hooker.hook()
    hooker.logging("log")
