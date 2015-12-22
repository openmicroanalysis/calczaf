import os
import shutil
import argparse

def rename(srcpath, recursive):
    if recursive and os.path.isdir(srcpath):
        for dirpath, dirnames, filenames in os.walk(srcpath):
            for name in filenames + dirnames:
                rename(os.path.join(dirpath, name), recursive)

    dirname, basename = os.path.split(srcpath)
    dstpath = os.path.join(dirname, basename.lower())
    if not os.path.exists(dstpath):
        shutil.move(srcpath, dstpath)
        print('Moved {0} to {1}'.format(srcpath, dstpath))

def main():
    description='Rename files and directories to lowercase'
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('paths', nargs='+', help='Path to files or directories')
    parser.add_argument('-r', '--recursive', action='store_true',
                        help='Recursively search directory')

    args = parser.parse_args()

    recursive = args.recursive
    for path in args.paths:
        rename(path, recursive)

if __name__ == '__main__':
    main()