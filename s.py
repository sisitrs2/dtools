import os
import argparse
from sys import platform
import getpass

IGNORE = ["AppData", "Application Data", "Cookies", "Intel", "IntelGraphicsProfiles", "Links", "MicrosoftEdgeBackups", "OneDrive", "AndroidStudioProjects"]

if platform == "win32":
    DEFAULTS = {"System32": "C:\\Windows\\System32"}
    USER = "C:\\Users\\" + getpass.getuser()
elif platform == "linux":
    DEFAULTS = {"System32": "/mnt/C/Windows/System32"}
    USER = "/mnt/C/Users/" + getpass.getuser()
else:
    print("OS Unidentified!!!")


def searchFile(filename, skip = False):
    try:
        files = os.listdir(".")
    except PermissionError:
        return False

    op_f = []

    for f in files:
        if skip and f == skip:
            skip == False
            continue
        if f.lower().startswith(filename.lower()):
            op_f.append(f)
            if not skip:
                return f

    if op_f != []:
        return op_f[0]

    return ""
        

def searchPath(d, path, origin = True, complete = False, all = False, nosearch = "", skip = False):

    if ".." in d or d[0] == ".":
        return d
    
    if d.lower() == path.split("\\")[-1].lower():
        if not skip:
            if all:
                print(path)
            else:
                return path
        elif d.lower() == skip.lower():
            skip = False

    # Search on current directory and sub.
    try:
        dirs = [name for name in os.listdir(path) if os.path.isdir(os.path.join(path, name)) and "." not in name and name != nosearch]
    except PermissionError:
        return False

    if not complete:
        if d.lower() in [dir.lower() for dir in dirs]:# and d.lower() != nosearch:
            if all:
                if path == os.getcwd():
                    print(os.path.join(path, d))
            else:
                return os.path.join(path, d)
    else:
        for dir in [dir.lower() for dir in dirs]:
            if dir.startswith(d.lower()):
                if not skip:
                    if all:
                        print(os.path.join(path, dir))
                    else:
                        return os.path.join(path, dir)
                elif dir.lower() == skip.lower():
                    skip = False

    for name in dirs:
        if name not in IGNORE and name != nosearch:
            p = searchPath(d, os.path.join(path, name), False, complete=complete, all=all, skip=skip)
            if p:
                if all:
                    print(p)
                else:
                    return p
    
    # Search on parent directory.
    if os.path.realpath(path) != USER and origin:
        path = path.split("\\")
        nos = path[-1]
        path = "\\".join(path[:-1])
        p = searchPath(d, path, origin=True, complete=complete, all=all, nosearch=nos, skip=skip)
        if p:
            if not skip:
                if all:
                    print(p)
                else:
                    return p
            elif p.lower() == skip.lower():
                skip = False


def DEFAULTS_val(key):
    for k, v in DEFAULTS.items():
        if k.lower() == key.lower():
            return v
    return False


def main():

    # Args parsing.
    parser = argparse.ArgumentParser()
    parser.add_argument("dir")
    parser.add_argument("-c", action="store_true", help="Complete path.")
    parser.add_argument("-a", action="store_true", help="Print all paths.")
    parser.add_argument("-f", action="store_true", help="Search files in current directory.")
    parser.add_argument("-s", help="Skip until after.")
    args = parser.parse_args()

    # Do search.
    if args.f:
        output = searchFile(args.dir, skip=args.s)
        if ' ' in output:
            output.split()
            output = '\\ '.join(output)
        print(output, end="")
        return
    else:
        output = searchPath(args.dir, os.getcwd(), complete=(True if args.c else False), all=(True if args.a else False), skip=args.s)
    if output:
        if ' ' in output:
            output = output.split()
            output = '\\ '.join(output)
        print(output, end="")

if __name__ == "__main__":
    main()

