import os
import argparse

RUN_FILE = 'rundll.c'

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('dllfile', help='Enter dll file.')
    parser.add_argument('-c', '--compile', action='store_const', const=compile, help='Flag for compile.')
    args = parser.parse_args()
    if not args.dllfile:
        print('Enter dll file.')
    
    exports = {}
    with open(args.dllfile, "r") as f:
        lines = f.read().splitlines()
        for l in lines:
            l = l.split(" ")
            if l[0] == "EXPORT":
                exports[l[2]] = l[1]

    files = os.listdir('.')
    dll = args.dllfile.split(".")[0]
    header = (dll + '.h')
    if header not in files:
        with open(header, 'a') as h:
            data = "#define EXPORT __declspec(dllexport)\n\n"
            for func, ret in exports.items():
                data += "EXPORT " + ret + " " + func + ";\n"
            data += "\n"
            h.write(data)

    if args.compile:
        os.system(f'gcc -c {args.dllfile}')
    dll_name = input('Enter dll name: ')

    if '.dll' not in dll_name:
        dll_name += '.dll'
    if args.compile:
        os.system(f'gcc -shared -o {dll_name} -Wl,--out-implib,libstdll.a {dll}.o')
    else:
        print('')
        print(f'gcc -c {args.dllfile}')
        print(f'gcc -shared -o {dll_name} -Wl,--out-implib,libstdll.a {dll}.o')

    if RUN_FILE in files:
        os.remove(RUN_FILE)
    
    with open(RUN_FILE, 'a') as r:
        data = ""
        data += f'#include "{header}"\n\n'
        data += 'int main()\n{\n'
        for func, ret in exports.items():
            if ret == 'void':
                data += f'   {func};\n'
        data += '   return 0;\n}\n'
        r.write(data)

    if args.compile:
        dll_name = dll_name.split('.')[0]
        os.system(f'gcc -o runDll rundll.c -L. -l{dll_name}')
    else:
        print(f'gcc -o runDll rundll.c -L. -l{dll_name}')

    if args.compile:
        os.remove(f'{dll}.o')
        os.remove('libstdll.a')


if __name__ == '__main__':
    main()
