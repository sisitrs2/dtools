import os
import argparse
from sys import platform

IGNORE = ["AppData", "Application Data", "Cookies", "Intel", "IntelGraphicsProfiles", "Links", "MicrosoftEdgeBackups", "OneDrive"]

if platform == "win32":
    DEFAULTS = {"System32": "C:\\Windows\\System32"}
    USERS = "C:\\Users"
elif platform == "linux":
    DEFAULTS = {"System32": "/mnt/C/Windows/System32"}
    USERS = "/mnt/C/Users"
else:
    print("OS Unidentified!!!")

def print_usage():
    print("This is help output")


def searchPath(d, path, origin = True):

    if ".." in d or d[0] == ".":
        return d

    # Search on current directory and sub.
    try:
        dirs = [name for name in os.listdir(path) if os.path.isdir(os.path.join(path, name)) and name[0] != "."]
    except PermissionError:
        return False
    if d.lower() in [dir.lower() for dir in dirs]:
        return os.path.join(path, d)
    else:
        for name in dirs:
            if name not in IGNORE:
                p = searchPath(d, os.path.join(path, name), False)
                if p:
                    return p
    
    # Search on parent directory.
    if origin and os.listdir(USERS) != os.listdir(path):
        p = searchPath(d, os.path.join("..", path))
        if p:
            return p


def DEFAULTS_val(key):
    for k, v in DEFAULTS.items():
        if k.lower() == key.lower():
            return v
    return False


def console(path):
    last = []
    last.append(os.getcwd())
    os.chdir(path)
    while True:
        print(os.getcwd() + ">", end="")
        try:
            cmd = input()
        except KeyboardInterrupt:
            print()
            continue
        if len(cmd) <= 1:
            continue
        elif cmd[0] == 'c' and cmd[1] == ' ':
            cmd = cmd.split(" ")
            if not cmd[1]:
                print("Help Page")
            elif cmd[1][0] == '-':
                if cmd[1][1] == 'e':
                    return
                elif cmd[1][1] == 'b':
                    if last:
                        p = last.pop()
                        os.chdir(p)
                    else:
                        return
            else:
                if cmd[1].lower() in [name.lower() for name in DEFAULTS.keys()]:
                    last.append(os.getcwd())
                    os.chdir(DEFAULTS_val(cmd[1]))
                else:
                    p = searchPath(cmd[1], ".")
                    if p:
                        last.append(os.getcwd())
                        os.chdir(p)
        elif cmd[0] == 'c' and cmd[1] == 'd' and cmd[2] == ' ':
            cmd = cmd.split(" ")
            if os.path.isdir(cmd[1]):
                last.append(os.getcwd())
                os.chdir(cmd[1])
            else:
                print("The system cannot find the path specified.\n")
        else:
            print()
            try:
                os.system(cmd)
            except KeyboardInterrupt:
                print()
                continue
            print()


def main():

    # Args parsing.
    parser = argparse.ArgumentParser()
    parser.add_argument("dir")
    args = parser.parse_args()

    # Action by flags.
    # TODO

    # Do search.
    if args.dir:
        if args.dir.lower() in [name.lower() for name in DEFAULTS.keys()]:
            console(DEFAULTS_val(args.dir))
        else:
            p = searchPath(args.dir, ".")
            if p:
                console(p)
            else:
                print("Directory was not found")
    else:
        print("Open Console.")
    

if __name__ == "__main__":
    main()

