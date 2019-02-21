import os
import shutil
import argparse


def rename(source_path, recursive):
    if recursive and os.path.isdir(source_path):
        for dir_path, dir_names, filenames in os.walk(source_path):
            for name in filenames + dir_names:
                rename(os.path.join(dir_path, name), recursive)

    dirname, basename = os.path.split(source_path)
    destination_path = os.path.join(dirname, basename.lower())
    if not os.path.exists(destination_path):
        shutil.move(source_path, destination_path)
        print('Moved {0} to {1}'.format(source_path, destination_path))


def main():
    description = 'Rename files and directories to lowercase'
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
