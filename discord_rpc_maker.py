import asyncio
import os
import sys
import threading
import webbrowser
from os import chdir
from time import sleep

import psutil
import PySimpleGUI as sg
import win32gui
import win32process
from pypresence import Presence
from validators import url

if getattr(sys, 'frozen', False): # Running as compiled
    chdir(sys._MEIPASS) # change current working directory to sys._MEIPASS

def ram_cpu():
    """
    Get CPU and RAM usage and return them as a tuple
    :return: A tuple of strings.
    """
    cpu_per = round(psutil.cpu_percent(), 1) # Get CPU Usage
    mem_per = round(psutil.virtual_memory().percent, 1)
    ram = f"RAM : {str(mem_per)}%"
    cpu = f"CPU : {str(cpu_per)}%"
    return ram, cpu


def setup_rpc(app_id):
    """
    It sets up the RPC connection to the Discord client.

    :param app_id: The application ID from Discord's developer portal
    :return: The RPC object.
    """
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    RPC = Presence(app_id, pipe=0)  # Initialize the client class
    RPC.connect()
    return RPC


def set_rpc(app_id): #TODO: make it compatible with presets
    """
    It sets the RPC to the values specified in the GUI.

    :param app_id: The ID of the application you want to use to update your status. The application ID from Discord's developer portal.
    """
    table = {
        "Details:": "details",
        "State:": "state",
        "Party Size: < int > , < int > ": "party_size",
        "Large Image Key:": "large_image",
        "Large Image Text:": "large_text",
        "Small Image Key:": "small_image",
        "Small Image Text:": "small_text"
    }
    RPC = setup_rpc(app_id)
    thread = threading.current_thread()
    while thread.do_run == True:
        buttons_dict = {"1": {"label": values["-BUTTON1_LABEL-"], "url": values["-BUTTON1_URL-"]}, "2": {"label": values["-BUTTON2_LABEL-"], "url": values["-BUTTON2_URL-"]}}
        buttons = []
        for element, value in buttons_dict.items():
            if (buttons_dict[element]["label"] == "" and buttons_dict[element]["url"] != "" and url(buttons_dict[element]["url"]) != True) or (buttons_dict[element]["label"] != "" and buttons_dict[element]["url"] == "") or (buttons_dict[element]["label"] != "" and buttons_dict[element]["url"] != "" and url(buttons_dict[element]["url"]) != True):
                queue.append("""sg.popup_error("Label and URL need to be both specified and URL needs to be valid!", title="Invalid Button or URL format.")""")
            elif buttons_dict[element]["label"] != "" and buttons_dict[element]["url"] != "" and url(buttons_dict[element]["url"]) == True:
                buttons.append({"label": buttons_dict[element]["label"], "url": value["url"]})

        fields_clean = {}
        for key, value in zip(fields.keys(), values):
            if key in ["Application ID*:"] or key.startswith("Button") or values[value] == "":
                continue
            elif key == "Party Size: < int > , < int > " and values[value] != "" and values['-STATE-'] == "":
                queue.append("""sg.popup_error("State needs to be specified if Party Size is specified!", title="Party Size and State")""")
            elif key == "Party Size: < int > , < int > " and values[value] != "":
                v = values[value].split(", ")
                if len(v) != 2:
                    queue.append("""sg.popup_error("Party Size needs to be in the format of < int > , < int > ", title="Invalid Party Size format.")""")
                    continue
                try:
                    v = (int(v[0]), int(v[1]))
                except ValueError:
                    queue.append("""sg.popup_error("Party Size needs to be in the format of < int > , < int > ", title="Invalid Party Size format.")""")
                if min(v) > 0:
                    fields_clean[table[key]] = v
                else:
                    queue.append("""sg.popup_error("Party Size needs to be bigger than 0 for both integers", title="Invalid Party Size value.")""")
            elif key in ["Large Image Text:", "Small Image Text:", "Details:", "State:"]:
                if len(values[value]) >= 2:
                    fields_clean[table[key]] = values[value]
                else:
                    queue.append(f"""sg.popup_error("The text must be at least 2 characters long.", title="{key}")""")
            else:
                fields_clean[table[key]] = values[value]
        params = ", ".join(f"{key} = '{value}'" if key != "party_size" else f"party_size={fields_clean['party_size']}" for key, value in fields_clean.items())
        if params != "":
            if buttons != []:
                params = f"{params}, buttons={buttons}"
        elif buttons != []:
            params = f"buttons={buttons}"
        cmd = f"RPC.update({params})"
        exec(cmd)
        for _ in range(15):
            if thread.do_run == False:
                break
            sleep(1)
    RPC.close()


def set_rpc_ram_cpu(app_id):
    """
    This function sets up a RPC connection to the app_id provided.
    Then it starts a thread that will run until the thread.do_run variable is set to False.
    The thread will update the RPC connection with the current RAM and CPU usage every 15 seconds.
    The thread will close the RPC connection when the thread is stopped

    :param app_id: The ID of the application you want to use to update your status. The application ID from Discord's developer portal.
    """
    RPC = setup_rpc(app_id)
    thread = threading.current_thread()
    while thread.do_run == True:
        ram, cpu = ram_cpu()
        RPC.update(details=ram, state=cpu)
        for _ in range(15):
            if thread.do_run == False:
                break
            sleep(1)
    RPC.close()


def parent_pid_process(pid):
    try:
        return (
            psutil.Process(pid)
            if psutil.Process(pid).parent() is None
            else psutil.Process(pid).parent()
        )
    except psutil.NoSuchProcess:
        return psutil.Process(pid)

def set_rpc_current_window(app_id):
    RPC = setup_rpc(app_id)
    thread = threading.current_thread()
    while thread.do_run == True:
        #* find the parent pid of the active window
        window = win32gui.GetWindowText(win32gui.GetForegroundWindow()) # Get the name of the current active window
        hwnd = win32gui.FindWindow(None, window) # Get the handle of the current active window
        _, pid = win32process.GetWindowThreadProcessId(hwnd) # Get the process ID of the current active window

        # get Process Object of the higher parent of the current active window if none return Process object of pid
        parent_process = parent_pid_process(pid)

        #* get cpu usage of the active window and its children
        #* get ram usage of the active window and its children
        cpu = parent_process.cpu_percent(0.1) / psutil.cpu_count()
        ram = parent_process.memory_full_info().uss
        for child in parent_process.children(recursive=True):
            with child.oneshot():
                try:
                    cpu += child.cpu_percent(0.1) / psutil.cpu_count()
                    ram += child.memory_full_info().uss
                except:pass
        ram = ram / (1024 ** 2)

        RPC.update(details=window, state=f"cpu: {cpu:.2f}% | ram: {ram:.0f}MB/{psutil.virtual_memory().total / (1024 ** 2):.0f}MB")
        for _ in range(15):
            if thread.do_run == False:
                break
            sleep(1)
    RPC.close()

fields = {
    "Application ID*:": {"key": "-APP_ID-", "tooltip": "Your Discord application ID. https://discord.com/developers/applications"},
    "Details:": {"key": "-DETAILS-", "tooltip": "The details of your presence."},
    "State:": {"key": "-STATE-", "tooltip": "The state of your presence."},
    "Party Size: < int > , < int > ": {"key": "-PARTY_SIZE-", "tooltip": "The size of your party."},
    "Large Image Key:": {"key": "-LARGE_IMAGE_KEY-", "tooltip": "The key of the large image."},
    "Large Image Text:": {"key": "-LARGE_IMAGE_TEXT-", "tooltip": "The text of the large image."},
    "Small Image Key:": {"key": "-SMALL_IMAGE_KEY-", "tooltip": "The key of the small image."},
    "Small Image Text:": {"key": "-SMALL_IMAGE_TEXT-", "tooltip": "The text of the small image."},
    "Button 1 Label:": {"key": "-BUTTON1_LABEL-", "tooltip": "The label of the first button."},
    "Button 1 URL:": {"key": "-BUTTON1_URL-", "tooltip": "The URL of the first button."},
    "Button 2 Label:": {"key": "-BUTTON2_LABEL-", "tooltip": "The label of the second button."},
    "Button 2 URL:": {"key": "-BUTTON2_URL-", "tooltip": "The URL of the second button."}
}

presets = {
    "": "",
    "RAM/CPU": set_rpc_ram_cpu,
    "Current Activity": set_rpc_current_window
}

sg.theme('DarkAmber')
layout = [[sg.Text("* fields are mandatory | RPC is updated every 15 seconds (ratelimit)")]]

layout.extend([sg.Text(text, size=(20, 1)), sg.Input(key=key["key"], tooltip=key["tooltip"])] for text, key in fields.items())

layout.append(
    [
        [
            sg.Text('Presets:'),
            sg.Combo(tuple(presets), key='-PRESETS-', enable_events=True)
        ],
        [
            sg.Button("ON", key="-SWITCH-", button_color="green"),
            sg.Button("Clear", key="-CLEAR-", button_color="grey"),
        ],
    ]
)


window = sg.Window('Discord RPC Maker', layout=layout, icon="MMA.ico")


class QueueEvents():
    """
    This class is a helper class that will be used to pass PySimpleGUI events that need to be run in main thread.

    The class has two methods:

    1. append: This method will append the events that need to be run in main thread.
    2. pop: This method will return the first event in the queue to be executed and removes it from the queue.
    """

    def __init__(self):
        self.queue = []

    def append(self, other) -> list:
        self.queue.append(other)
        return self.queue

    def pop(self) -> str:
        if self.queue:
            return self.queue.pop(0)

queue = QueueEvents()

if __name__ == "__main__":
    while True:
        event, values = window.Read(timeout=100)
        if event is None or event in [sg.WIN_CLOSED]: #TODO: detect if thread is running and set thread.do_run = False
            os._exit(0)
        elif queue.queue:
            exec(queue.pop())
        elif event == "-SWITCH-" and values["-APP_ID-"].isdecimal() and len(values["-APP_ID-"]) == 18 and window["-SWITCH-"].ButtonText == 'ON' and values["-PRESETS-"] == "":
            thread = threading.Thread(target=set_rpc, args=(values["-APP_ID-"], ))
            thread.do_run = True
            thread.start()
            window["-SWITCH-"].Update("OFF", button_color="red")
        elif event == "-SWITCH-" and values["-APP_ID-"].isdecimal() and len(values["-APP_ID-"]) == 18 and window["-SWITCH-"].ButtonText == 'ON' and values["-PRESETS-"] != "":
            thread = threading.Thread(target=presets[values["-PRESETS-"]], args=(values["-APP_ID-"], ))
            thread.do_run = True
            thread.start()
            window["-SWITCH-"].Update("OFF", button_color="red")
        elif event == "-PRESETS-" and values["-APP_ID-"].isdecimal() and len(values["-APP_ID-"]) == 18 and window["-SWITCH-"].ButtonText == 'OFF' and values["-PRESETS-"] == "":
            thread.do_run = False
            thread = threading.Thread(target=set_rpc, args=(values["-APP_ID-"], ))
            thread.do_run = True
            thread.start()
        elif event == "-PRESETS-" and values["-APP_ID-"].isdecimal() and len(values["-APP_ID-"]) == 18 and window["-SWITCH-"].ButtonText == 'OFF' and values["-PRESETS-"] != "":
            thread.do_run = False
            thread = threading.Thread(target=presets[values["-PRESETS-"]], args=(values["-APP_ID-"], ))
            thread.do_run = True
            thread.start()
        elif event == "-SWITCH-" and window["-SWITCH-"].ButtonText == 'ON' and not values["-APP_ID-"].isdecimal() and len(values["-APP_ID-"]) != 18:
            response = sg.popup_yes_no("Application ID must be 18 digits long. Would you like to be redirected to the discord webpage to get your application ID?")
            if response == "Yes":
                webbrowser.open("https://discord.com/developers/applications")
        elif event == "-SWITCH-" and window["-SWITCH-"].ButtonText == 'OFF':
            thread.do_run = False
            window["-SWITCH-"].Update("ON", button_color="green")
        elif event == "-CLEAR-":
            for element in values:
                if element not in ["-APP_ID-"]:
                    window[element].Update('')
    window.close()
