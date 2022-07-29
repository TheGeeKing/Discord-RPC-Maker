import asyncio
import configparser
import contextlib
import json
import os
import sys
import threading
import time
import webbrowser
from os import chdir
from time import sleep
from typing import Callable

import psutil
import PySimpleGUI as sg
import win32gui
import win32process
from pypresence import Presence
from validators import url
from win32com.client import Dispatch


class QueueEvents(): #TODO: use python queue and update put method to check if already in queue
    """
    This class is a helper class that will be used to pass PySimpleGUI events that need to be run in main thread.

    The class has two methods:

    1. append: This method will append the events that need to be run in main thread.
    2. pop: This method will return the first event in the queue to be executed and removes it from the queue.
    """

    def __init__(self, max_size: int = 2):
        self.queue = []
        self.max_size = max_size # The number of popup display will be +1 due to the fact that the first one will be instantly displayed while 2 others will stay in the queue.

    def append(self, other) -> list[tuple]:
        if len(self.queue) == self.max_size:
            self.queue.pop(0)
        if other in self.queue:
            return self.queue
        self.queue.append(other)
        return self.queue

    def pop(self) -> tuple:
        if self.queue:
            return self.queue.pop(0)


class Config: #TODO: use python configparser and update the set and write method
    def __init__(self, path):
        self.path = path
        self.config = configparser.ConfigParser()
        self.read()

    def read(self):
        self.config.read(self.path)
        return self.config

    def write(self):
        with open(self.path, "w", encoding="utf-8") as configfile:
            self.config.write(configfile, space_around_delimiters=False)

    def has_section(self, section):
        return self.config.has_section(section)

    def has_option(self, section, option):
        return self.config.has_option(section, option)

    def add_section(self, section):
        self.config.add_section(section)
        self.write()

    def set(self, section, option, value):
        if not self.has_section(section):
            self.add_section(section)
        self.config.set(section, option, value)
        self.write()

    def get(self, section, option):
        return self.config.get(section, option)


def ram_cpu() -> tuple[str, str]:
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
                queue.append(("Label and URL need to be both specified and URL needs to be valid!", "Invalid Button or URL format."))
            elif buttons_dict[element]["label"] != "" and buttons_dict[element]["url"] != "" and url(buttons_dict[element]["url"]) == True:
                buttons.append({"label": buttons_dict[element]["label"], "url": value["url"]})

        fields_clean = {}
        for key, value in zip(fields.keys(), values):
            if key in ["Application ID*:"] or key.startswith("Button") or values[value] == "":
                continue
            elif key == "Party Size: < int > , < int > " and values[value] != "" and values['-STATE-'] == "":
                queue.append(("State needs to be specified if Party Size is specified!", "Party Size and State"))
            elif key == "Party Size: < int > , < int > " and values[value] != "":
                v:tuple[int, int] = values[value].split(", ")
                if len(v) != 2:
                    queue.append(("Party Size needs to be in the format of < int > , < int > ", "Invalid Party Size format."))
                    continue
                try:
                    v:tuple[int, int] = (int(v[0]), int(v[1]))
                except ValueError:
                    queue.append(("Party Size needs to be in the format of < int > , < int >", "Invalid Party Size format."))
                if min(v) > 0:
                    fields_clean[table[key]] = v
                else:
                    queue.append(("Party Size needs to be bigger than 0 for both integers", "Invalid Party Size value."))
            elif key in ["Large Image Text:", "Small Image Text:", "Details:", "State:"]:
                if len(values[value]) >= 2:
                    fields_clean[table[key]] = values[value]
                else:
                    queue.append(("The text must be at least 2 characters long.", key))
            else:
                fields_clean[table[key]] = values[value]
        params = ", ".join(f"{key}=\"{value}\"" if key != "party_size" else f"party_size={fields_clean['party_size']}" for key, value in fields_clean.items())
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


def parent_pid_process(pid) -> psutil.Process:
    """
    The parent_pid_process function returns the parent process of a given pid. If no parent is found, it returns Process object of given pid.

    :param pid: Specify the process id
    :return: The parent process of the given pid
    """
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
        parent_process: psutil.Process = parent_pid_process(pid)

        #* get cpu usage of the active window and its children
        #* get ram usage of the active window and its children
        cpu = parent_process.cpu_percent(0.1) / psutil.cpu_count()
        ram = parent_process.memory_full_info().uss
        for child in parent_process.children(recursive=True):
            with child.oneshot():
                with contextlib.suppress(Exception):
                    cpu += child.cpu_percent(0.1) / psutil.cpu_count()
                    ram += child.memory_full_info().uss
        ram = ram / (1024 ** 2)

        RPC.update(details=window, state=f"cpu: {cpu:.2f}% | ram: {ram:.0f}MB/{psutil.virtual_memory().total / (1024 ** 2):.0f}MB")
        for _ in range(15):
            if thread.do_run == False:
                break
            sleep(1)
    RPC.close()


def is_valid_preset_file(file_: str, output_error: bool = False) -> bool:
    # verify that the file is valid and has the all the keys in the preset_format
    if not file_.endswith('.json') or not os.path.isfile(os.path.join(presets_dir, file_)):
        return False
    if file_ != "Example.json":
        with open(os.path.join(presets_dir, file_), 'r') as f:
            try:
                data = json.load(f)
                if all(key in data for key in preset_format):
                    return True
                if output_error:
                    x = '\n'.join(preset_format) # due to the fact that: SyntaxError: f-string expression part cannot include a backslash. When f"{'\n'join()}""
                    queue.append(f"{file_} is not a valid preset file.\nFormat:\n{x}\nare required even if empty.")
                return False
            except Exception as e:
                if output_error: queue.append(f"{file_} is not a valid preset file. It is not a valid JSON file.\n{e}")
                return False


def build_presets_dropdown() -> dict[str, Callable]:
    presets: dict[str, Callable] = {
    "": "",
    "RAM/CPU": set_rpc_ram_cpu,
    "Current Activity": set_rpc_current_window,
    "Example": set_rpc
    }

    preset_filename = [file_ for file_ in os.listdir(presets_dir) if is_valid_preset_file(file_, output_error=True)]

    for preset in preset_filename:
        presets[preset[:-5]] = set_rpc

    return presets


def update_presets_dropdown() -> dict[str, Callable]:
    presets: dict[str, Callable] = build_presets_dropdown()
    window['-PRESETS-'].update(values=tuple(presets))
    return presets


def clear_fields():
    """
    The clear_fields function clears all the fields in the form except -APP_ID-.
    It is used to quickly empty fields so user can write a new RPC Status.
    """
    for element in values:
        if element not in ["-APP_ID-"]:
            window[element].Update('')


def create_shortcut(path, target="", w_dir="", icon="", arguments=""):
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.Arguments = f"\"{arguments}\""
    shortcut.WorkingDirectory = w_dir
    shortcut.WindowStyle = 7
    if icon != "":
        shortcut.IconLocation = icon
    shortcut.save()


def is_process_running(process_running):
    """
    Check if there is any running process that contains the given name process_running.
    """
    #Iterate over the all the running process
    for proc in psutil.process_iter():
        with contextlib.suppress(psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # Check if process name contains the given name string.
            if process_running.lower() in proc.name().lower():
                return True
    return False


if __name__ == "__main__":

    if getattr(sys, "frozen", False): # Running as compiled
        chdir(sys._MEIPASS) # change current working directory to sys._MEIPASS

    if len(sys.argv) > 1: # if it starts from start-up
        while not is_process_running("discord.exe"):
            time.sleep(10) # check if discord is running every 10 seconds
        sleep(10) # wait for discord to load, so RPC is not runned while discord is loading and not ready yet

    """This part creates if it doesn't exist the needed directories and make sure the Example.json file exists in the user's home directory."""
    temp_dir = os.path.dirname(os.path.abspath(os.path.realpath(__file__)))
    file_path = os.path.abspath(os.path.realpath(sys.argv[0])) # file path of the current file
    file_dir = os.path.dirname(file_path) # directory of the current file
    appdata_dir = os.getenv("APPDATA")
    startup_dir = os.path.join(appdata_dir, r"Microsoft\Windows\Start Menu\Programs\Startup")
    home_dir = os.path.expanduser("~")
    thegeeking_dir = os.path.join(home_dir, ".TheGeeKing")
    discord_rpc_maker_dir = os.path.join(thegeeking_dir, "Discord-RPC-Maker")
    presets_dir = os.path.join(discord_rpc_maker_dir, "presets")
    # create a directory called .TheGeeKing if it doesn't exist
    if not os.path.exists(thegeeking_dir): os.mkdir(thegeeking_dir)
    # create a directory called Discord-RPC-Maker if it doesn't exist
    if not os.path.exists(discord_rpc_maker_dir): os.mkdir(discord_rpc_maker_dir)
    # create a presets folder if it doesn't exist
    if not os.path.exists(presets_dir): os.mkdir(presets_dir)
    # if config.ini doesn't exist, create it and write the default values
    if not os.path.exists(os.path.join(discord_rpc_maker_dir, "config.ini")):
        with open(os.path.join(discord_rpc_maker_dir, "config.ini"), "w", encoding="utf-8") as f:
            f.write("")
    config = Config(os.path.join(discord_rpc_maker_dir, "config.ini"))
    if not config.has_section("config"):
        config.set("config", "start-up", "")
    elif not config.has_option("config", "start-up"):
        config.set("config", "start-up", "")

    queue = QueueEvents()

    fields: dict[str, dict[str, str]] = {
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

    sg.theme("DarkAmber")
    layout = [[sg.Text("* fields are mandatory | RPC is updated every 15 seconds (ratelimit)")]]

    layout.extend([sg.Text(text, size=(20, 1)), sg.Input(key=key["key"], tooltip=key["tooltip"])] for text, key in fields.items())

    layout.append(
        [
            [
                sg.Text("Presets:"),
                sg.Combo("", key="-PRESETS-", enable_events=True, readonly=True, size=15, tooltip="Select a preset"),
                sg.Push(),
                sg.Input("", key="-INPUT_PRESET-", size=(20, 1), tooltip="Enter a preset name to save or delete it."),
                sg.Button("Save", key="-SAVE-", button_color="lightblue"),
                sg.Button("Delete", key="-DELETE-", button_color="red")
            ],
            [
                sg.Button("ON", key="-SWITCH-", button_color="green"),
                sg.Button("Clear", key="-CLEAR-", button_color="grey"),
                sg.Push(),
                sg.Button("Remove auto start-up", key="-AUTO_STARTUP-", tooltip="Remove the auto start-up shortcut.\n(%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup)", button_color="grey", visible=False)
            ],
        ]
    )


    window = sg.Window("Discord RPC Maker", layout=layout, icon="MMA.ico", finalize=True)

    preset_format = list(fields)
    elements_key = [value["key"] for value in fields.values()]

    presets: dict[str, Callable] = update_presets_dropdown()

    """If it starts from start-up"""
    auto_start = 0
    if len(sys.argv) > 1:
        if config.get("config", "start-up") == "" or not is_valid_preset_file(config.get("config", "start-up")): # auto start-up is invalid, shortcut and config is reset
            if os.path.exists(os.path.join(startup_dir, "Discord RPC Maker - Startup.lnk")):
                os.remove(os.path.join(startup_dir, "Discord RPC Maker - Startup.lnk"))
            config.set("config", "start-up", "")
        else: # auto start-up is valid
            window["-PRESETS-"].Update(config.get("config", "start-up"))
            window.write_event_value("-PRESETS-", config.get("config", "start-up"))
            auto_start = 1
            window["-AUTO_STARTUP-"].Update(visible=True)
            window.minimize()
    elif config.get("config", "start-up") != "" and os.path.exists(os.path.join(startup_dir, "Discord RPC Maker - Startup.lnk")): #If manually started by user, check if the config.ini start-up isn't empty and that the startup shortcut exists and if so display the button to remove it and fill the fields with the preset selected as start-up.
        if not is_valid_preset_file(config.get("config", "start-up")): # start-up is invalid, shortcut and config is reset
            if os.path.exists(os.path.join(startup_dir, "Discord RPC Maker - Startup.lnk")):
                os.remove(os.path.join(startup_dir, "Discord RPC Maker - Startup.lnk"))
            config.set("config", "start-up", "")
        else: # start-up is valid
            window["-PRESETS-"].Update(config.get("config", "start-up"))
            window.write_event_value("-PRESETS-", config.get("config", "start-up"))
            window["-AUTO_STARTUP-"].Update(visible=True)
    elif config.get("config", "start-up") != "" and not os.path.exists(os.path.join(startup_dir, "Discord RPC Maker - Startup.lnk")):
        config.set("config", "start-up", "")

    while True:
        event, values = window.Read(timeout=100) # read every 1/10 second (100ms) to update the RPC if changed while running
        #TODO: statements might need refactoring so thread takes directly the values and no need to stop them and restart them here
        if event is None or event in [sg.WIN_CLOSED]: #TODO: detect if thread is running and set thread.do_run = False
            window.close()
            os._exit(0)
        elif queue.queue: # if a thread added something to the queue it will show it here in a popup_error
            """queue.pop() should return a tuple containing as 0 the message and as 1 the title.
            if the len of the tuple returned by queue.pop() == 1 then the title is by default 'Error'"""
            queue_pop = queue.pop()
            sg.popup_error(queue_pop[0], title="Error" if len(queue_pop) == 1 else queue_pop[1], icon="MMA.ico")
        elif event == "-SWITCH-" and values["-APP_ID-"].isdecimal() and len(values["-APP_ID-"]) == 18 and window["-SWITCH-"].ButtonText == 'ON' and values["-PRESETS-"] == "": # standard RPC with no preset
            thread = threading.Thread(target=set_rpc, args=(values["-APP_ID-"], ))
            thread.do_run = True
            thread.start()
            window["-SWITCH-"].Update("OFF", button_color="red")
        elif event == "-SWITCH-" and values["-APP_ID-"].isdecimal() and len(values["-APP_ID-"]) == 18 and window["-SWITCH-"].ButtonText == 'ON' and values["-PRESETS-"] != "": # RPC with standard preset
            thread = threading.Thread(target=presets[values["-PRESETS-"]], args=(values["-APP_ID-"], ))
            thread.do_run = True
            thread.start()
            window["-SWITCH-"].Update("OFF", button_color="red")
        elif event == "-PRESETS-" and values["-APP_ID-"].isdecimal() and len(values["-APP_ID-"]) == 18 and window["-SWITCH-"].ButtonText == 'OFF' and values["-PRESETS-"] == "":
            thread.do_run = False
            thread = threading.Thread(target=set_rpc, args=(values["-APP_ID-"], ))
            thread.do_run = True
            thread.start()
        elif event == "-PRESETS-" and values["-APP_ID-"].isdecimal() and len(values["-APP_ID-"]) == 18 and window["-SWITCH-"].ButtonText == 'OFF' and values["-PRESETS-"] != "" and presets[values["-PRESETS-"]] != set_rpc:
            thread.do_run = False
            thread = threading.Thread(target=presets[values["-PRESETS-"]], args=(values["-APP_ID-"], ))
            thread.do_run = True
            thread.start()
        elif event == "-PRESETS-" and values["-PRESETS-"] != "" and presets[values["-PRESETS-"]] == set_rpc:
            with open(os.path.join(presets_dir, f"{values['-PRESETS-']}.json") if values["-PRESETS-"] != "Example" else r"config\Example.json", 'r', encoding='utf-8') as f:
                data = json.load(f)
            for key, key_ in zip(elements_key, data.keys()):
                window[key].Update(data[key_])
            window["-PRESETS-"].Update(data["Preset"])
            if auto_start:
                window["-SWITCH-"].click()
                auto_start = 0
        elif event == "-SWITCH-" and window["-SWITCH-"].ButtonText == 'ON' and not values["-APP_ID-"].isdecimal() and len(values["-APP_ID-"]) != 18:
            response = sg.popup_yes_no("Application ID must be 18 digits long. Would you like to be redirected to the discord webpage to get your application ID?", icon="MMA.ico")
            if response == "Yes":
                webbrowser.open("https://discord.com/developers/applications")
        elif event == "-SWITCH-" and window["-SWITCH-"].ButtonText == 'OFF':
            thread.do_run = False
            window["-SWITCH-"].Update("ON", button_color="green")
        elif event == "-CLEAR-":
            clear_fields()
        elif event == "-SAVE-":
            if values["-INPUT_PRESET-"] in ["Example", "RAM/CPU", "Current Activity"]:
                sg.popup_ok(f"You can't save a preset with the name \'{values['-INPUT_PRESET-']}\'.", icon="MMA.ico")
            elif values["-INPUT_PRESET-"] != "":
                try:
                    v: dict[str, str] = {key: values[key_] for key, key_ in zip(preset_format, elements_key)}
                    v["Preset"] = values["-PRESETS-"]
                    answer = sg.popup_yes_no("Do you want this preset to be executed at start-up?", icon="MMA.ico")
                    if answer == "Yes":
                        create_shortcut(os.path.join(startup_dir, "Discord RPC Maker - Startup.lnk"), file_path, file_dir, os.path.join(temp_dir, "MMA.ico"), "--start-up")
                        window["-AUTO_STARTUP-"].Update(visible=True)
                        config.set("config", "start-up", values["-INPUT_PRESET-"])
                    with open(os.path.join(presets_dir, f"{values['-INPUT_PRESET-']}.json"), 'w', encoding='utf-8') as f:
                        json.dump(v, f, indent=4)

                    sg.popup_ok("Preset saved!", icon="MMA.ico")
                except Exception as e:
                    sg.popup_error(f"{e}", title="Error", icon="MMA.ico")
                window["-INPUT_PRESET-"].Update("")

                presets: dict[str, Callable] = update_presets_dropdown()
            else:
                sg.popup_error("Please enter a preset name.", title="Error", icon="MMA.ico")
        elif event == "-DELETE-":
            if values["-INPUT_PRESET-"] == "Example":
                sg.popup_ok("You cannot delete the preset 'Example'.", icon="MMA.ico")
            elif values["-INPUT_PRESET-"] != "":
                try:
                    os.remove(os.path.join(presets_dir, f"{values['-INPUT_PRESET-']}.json"))
                    sg.popup_ok("Successfully deleted preset.", title="Success", icon="MMA.ico")
                except Exception as e:
                    sg.popup_error(f"{e}", title="Error", icon="MMA.ico")
                window["-INPUT_PRESET-"].Update("")

                presets: dict[str, Callable] = update_presets_dropdown()
            else:
                sg.popup_error("Please enter a preset name.", title="Error", icon="MMA.ico")
        elif event == "-AUTO_STARTUP-":
            if os.path.exists(os.path.join(startup_dir, "Discord RPC Maker - Startup.lnk")):
                os.remove(os.path.join(startup_dir, "Discord RPC Maker - Startup.lnk"))
                config.set("config", "start-up", "")
                window["-AUTO_STARTUP-"].Update(visible=False)
                sg.popup_ok("Successfully removed startup shortcut.", title="Success", icon="MMA.ico")
