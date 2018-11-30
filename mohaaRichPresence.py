from pypresence import Presence
import win32con
import win32gui
import socket
import time

def readMOHAA():
    try:
        print("Looking for MOHAA.")
        # Find the MOH console window
        hwnd = win32gui.FindWindow("MOHAA WinConsole", None)
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        clearBtn = win32gui.FindWindowEx(hwnd, None, "Button", "clear")
        # Click the clear button
        win32gui.SendMessage(hwnd, win32con.WM_COMMAND, 1, clearBtn)

        #Set cmdline with serverinfo
        cmdline = win32gui.GetDlgItem(hwnd, 101)
        win32gui.SendMessage(cmdline, win32con.WM_SETTEXT, None, "clientinfo")
        win32gui.PostMessage(cmdline, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)

        #Get console info
        console = win32gui.GetDlgItem(hwnd, 100)
        buffer = win32gui.PyMakeBuffer(25555)
        length = win32gui.SendMessage(console, win32con.WM_GETTEXT, 25555, buffer)
        result = buffer[0:length * 2].tobytes().decode("UTF-16")

        # Click the clear button
        win32gui.SendMessage(hwnd, win32con.WM_COMMAND, 1, clearBtn)

        # split on newline
        splitted = result.splitlines()

        for index, value in enumerate(splitted):
            if "Server: " in value:
                print("Found server! " + value[8:])
                return value[8:]

    except Exception as e:
        print("Error[MOHAA]: " + str(e))

def getServerDetails(server):
    try:
        ipport = server.split(':')
        IP = ipport[0]
        port = int(ipport[1])
        cmd = "getstatus"
        serverinfo = []

        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(5)

        sock.connect((IP, port))
        sock.send(b"\xFF\xFF\xFF\xFF\x02" + str.encode(cmd))
        received = sock.recv(65565)
        r = received.decode("ISO-8859-1").split("\\")
        servernameindex = r.index("sv_hostname")
        mapindex = r.index("mapname")

        serverinfo.append(r[servernameindex+1])
        serverinfo.append(r[mapindex+1])
        print("Translating IP to servername: " + r[servernameindex+1])
        sock.close()
        return serverinfo

    except Exception as e:
        print("Error[network]: " + str(e))
        print("Connection timeout to server.")
        sock.close()

def discordRP():
    try:
        client_id = '517748775792410635'
        RPC = Presence(client_id)
        print("Connecting to Discord Rich Presence.")
        RPC.connect()
        print("Succesfully connnected!")
    except Exception as e:
        print("Couldn't connect to Discord Rich Presence, error: \n" + str(e))

    while True:  # The presence will stay on as long as the program is running
        try:
            mohaaserver = readMOHAA()
            serverinfo = getServerDetails(mohaaserver)
            RPC.update(state=serverinfo[1], details=serverinfo[0], large_image="moh_icon", small_image=serverinfo[1][:3], ) # Set the presence
        except Exception as e:
            print("Error[Discord]: " + str(e))
        time.sleep(15)

if __name__ == '__main__':
    discordRP()
