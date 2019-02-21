"""
Script to synchronize CalcZAF source code from Probe Software with GitHub using
Travis CI.
Steps:

  * Pull repository from GitHub
  * Download ZIP containing CalcZAF source code from Probe Software website
  * Parse version text file and compare for new changes
  * Extract files from ZIP
  * Add modified files, commit and push to GitHub
"""

# Standard library modules.
import logging
import os
import sys
import urllib.request
import urllib.parse
import urllib.error
import tempfile
import zipfile
import argparse
import shutil
import subprocess
import re
import datetime
import io

# Third party modules.

# Local modules.

# Globals and constants variables.
logger = logging.getLogger(__name__)

VERSION_FILENAME = 'VERSION.TXT'


def parse_version(fp):
    changes = {}
    tags = {}

    change_pattern = re.compile('^\d{1,2}/\d{1,2}/\d{1,2}\t')
    tag_pattern = re.compile('v. \d+')

    lines = []
    for line in fp:
        line = line.strip()
        if not line:
            continue

        dt = None
        if change_pattern.match(line):
            lines = []
            dt, line = line.split('\t', 1)
            try:
                dt = datetime.datetime.strptime(dt, '%m/%d/%y')
            except ValueError:
                continue
            changes[dt] = lines
            lines.append(line)
        elif tag_pattern.match(line):
            if dt is not None:
                tag_value = line[2:].split('\t', 1)[0].strip()
                if tag_value.endswith('.'):
                    tag_value = tag_value[:-1]
                tags[dt] = tag_value
        else:
            lines.append(line)

    return changes, tags


def pull(work_dir, repos_url):
    logger.info('Running git pull/clone...')
    if '.git' in os.listdir(work_dir):
        args = ['git', 'pull']
    else:
        args = ['git', 'clone', repos_url, work_dir]
    subprocess.check_call(args, cwd=work_dir)
    logger.info('Running git pull/clone... DONE')


def has_changes(work_dir):
    logger.info('Running git status...')
    args = ['git', 'status', '--porcelain']
    output = subprocess.check_output(args, cwd=work_dir, universal_newlines=True)
    logger.info('Running git status... DONE')
    return bool(output)


def commit(work_dir, message, tag, do_commit=True, push=True):
    logger.info('Running git add...')
    args = ['git', 'add', '.']
    if not do_commit:
        args.append('--dry-run')
    logger.debug(args)
    if do_commit:
        subprocess.check_call(args, cwd=work_dir)
    logger.info('Running git add... DONE')

    logger.info('Running git commit...')
    args = ['git', 'commit', '-m', message]
    if not do_commit:
        args.append('--dry-run')
    logger.debug(args)
    if do_commit:
        subprocess.check_call(args, cwd=work_dir)
    logger.info('Running git commit... DONE')

    has_tag = False
    if tag is not None and do_commit:
        logger.info('Running git tag...')
        args = ['git', 'tag', tag]
        subprocess.check_call(args, cwd=work_dir)
        logger.info('Running git tag... DONE')
        has_tag = True

    logger.info('Running git push...')
    args = ['git', 'push', '--all']
    if not push:
        args.append('--dry-run')
    subprocess.check_call(args, cwd=work_dir)

    if has_tag:
        args = ['git', 'push', '--tags']
        if not push:
            args.append('--dry-run')
        subprocess.check_call(args, cwd=work_dir)

    logger.info('Running git push... DONE')


def compare(filepath, work_dir):
    logger.info('Reading current version...')
    work_filepath = os.path.join(work_dir, VERSION_FILENAME.lower())
    with open(work_filepath, 'r', errors='ignore') as fp:
        old_changes, old_tags = parse_version(fp)

    logger.info('Reading current version... DONE')

    logger.info('Reading zip version...')
    with zipfile.ZipFile(filepath, 'r') as z:
        with io.TextIOWrapper(z.open(VERSION_FILENAME), errors='ignore') as fp:
            new_changes, new_tags = parse_version(fp)
    logger.info('Reading zip version... DONE')

    logger.info('Comparing versions...')

    new_changes = dict((dt, val) for dt, val in new_changes.items() if dt not in old_changes)
    new_tags = dict((dt, val) for dt, val in new_tags.items() if dt not in old_tags)

    new_tag = new_tags[max(new_tags)] if new_tags else None
    logger.info('Comparing versions... DONE')

    for new_change in new_changes:
        logger.debug(new_change)

    for new_tag in new_tags:
        logger.debug(new_tags)

    return new_changes, new_tag


def compare_remove_files(filepath, work_dir, no_commit):
    logger.info('Comparing files %s...', filepath)

    zipfile_list = set()
    with zipfile.ZipFile(filepath, 'r') as z:
        for info in z.infolist():
            # Always use lower case
            filename = info.filename.lower()
            zipfile_list.add(filename)
            # Create directory if needed
            dirname = os.path.join(work_dir, os.path.dirname(filename))
            os.makedirs(dirname, exist_ok=True)

    work_dir_list = set(os.listdir(work_dir))
    # Ignore these files
    # noinspection SpellCheckingInspection
    ignore_files = [".travis.yml", ".travis", ".git", ".gitignore", "readme.md", "license", ".idea"]
    for ignore_file in ignore_files:
        try:
            work_dir_list.remove(ignore_file)
        except KeyError as message:
            logging.warning(message)

    removed_files = work_dir_list - zipfile_list

    logger.info('Comparing files %s... DONE', filepath)

    if len(removed_files) > 0:
        for removed_file in removed_files:
            logger.info('Running git rm %s ...', removed_file)
            args = ['git', 'rm', removed_file]
            if no_commit:
                args.append('--dry-run')
            try:
                subprocess.check_call(args, cwd=work_dir)
            except subprocess.CalledProcessError as message:
                logging.warning(message)
            logger.info('Running git rm %s... DONE', removed_file)

        if not no_commit:
            message = "Remove files not in the CalcZAF source zip file."
            logger.info('Running git commit...')
            args = ['git', 'commit', '-m', message]
            subprocess.check_call(args, cwd=work_dir)
            logger.info('Running git commit... DONE')


def extract(filepath, work_dir):
    logger.info('Extracting %s...', filepath)
    with zipfile.ZipFile(filepath, 'r') as z:
        for info in z.infolist():
            # Always use lower case
            filename = info.filename.lower()

            # Create directory if needed
            dirname = os.path.join(work_dir, os.path.dirname(filename))
            os.makedirs(dirname, exist_ok=True)

            # Write filename
            destination_path = os.path.join(work_dir, filename)
            with z.open(info, 'r') as fi, open(destination_path, 'wb') as fo:
                fo.write(fi.read())

    logger.info('Extracting %s... DONE', filepath)


def create_commit_message(changes):
    logger.info('Create commit message...')

    message = []
    message += ['Auto-sync on %s' % datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S'), '']

    for dt in sorted(changes):
        message.append(dt.strftime('%Y/%m/%d'))
        for line in changes[dt]:
            message.append(' ' * 4 + line)

    logger.info('Create commit message... DONE')

    return '\n'.join(message)


def download(url):
    def reporthook(block_num, block_size, total_size):
        """
        From http://stackoverflow.com/questions/13881092/download-progressbar-for-python-3
        """
        read_so_far = block_num * block_size
        if total_size > 0:
            percent = read_so_far * 1e2 / total_size
            s = "\r%5.1f%% %*d / %d" % (
                percent, len(str(total_size)), read_so_far, total_size)
            sys.stderr.write(s)
            if read_so_far >= total_size:  # near the end
                sys.stderr.write("\n")
        else:  # total size is unknown
            sys.stderr.write("read %d\n" % (read_so_far,))

    logger.info('Downloading...')
    filepath, _headers = urllib.request.urlretrieve(url, reporthook=reporthook)
    logger.info('Downloading... DONE')

    return filepath


def setup(work_dir):
    user_work_dir = True
    if not work_dir:
        work_dir = tempfile.mkdtemp()
        user_work_dir = False
        logger.info('Created work directory: %s', work_dir)

    if not os.path.exists(work_dir):
        os.makedirs(work_dir)
        logger.info('Created work directory: %s', work_dir)

    return work_dir, user_work_dir


def teardown(work_dir, user_work_dir, zip_filepath, user_zip):
    if not user_work_dir:
        shutil.rmtree(work_dir)
        logger.info('Removed work directory: %s', work_dir)
    if not user_zip:
        os.remove(zip_filepath)


def main():
    description = 'Sync CalcZAF from Probe Software to openMicroanalysis GitHub'
    parser = argparse.ArgumentParser(description=description)

    parser.add_argument('-v', '--verbose', action='store_true', default=False,
                        help='Verbose mode')
    parser.add_argument('-w', '--work_dir', help='Working directory')
    parser.add_argument('-u', '--url',
                        default='http://probesoftware.com/download/CALCZAF_SOURCE-E2.ZIP',
                        help='Url to CalcZAF source zip file')
    parser.add_argument('--reposurl', default='git@github.com:openmicroanalysis/calczaf.git')
    parser.add_argument('--no-pull', action='store_true', default=False,
                        help='Do not pull in working directory')
    parser.add_argument('--no-push', action='store_true', default=False,
                        help='Do not push after commit')
    parser.add_argument('--no-commit', action='store_true', default=False,
                        help='Do not commit change')

    args = parser.parse_args()

    level = logging.DEBUG if args.verbose else logging.INFO
    logger.addHandler(logging.StreamHandler())
    logger.setLevel(level)

    no_pull = args.no_pull
    if no_pull:
        logger.info('Pull disabled')
    no_commit = args.no_commit
    if no_commit:
        logger.info('Commit disabled')
    no_push = args.no_push
    if no_push:
        logger.info('Push disabled')

    work_dir = args.work_dir
    work_dir, user_work_dir = setup(work_dir)

    # Pull repository
    repos_url = args.reposurl
    if not no_pull:
        pull(work_dir, repos_url)

    # Download zip
    url = args.url
    try:
        zip_filepath = download(url)
        user_zip = False
    except urllib.error.URLError:
        zip_filepath = url
        user_zip = True
    except ValueError:
        zip_filepath = url
        user_zip = True

    # Compare versions
    version_changes, tag = compare(zip_filepath, work_dir)
    logger.info('%i new changes found in VERSION.TXT', len(version_changes))

    # Compare files
    compare_remove_files(zip_filepath, work_dir, no_commit)

    has_repos_changes = has_changes(work_dir)
    if has_repos_changes:
        logger.info('has changes found in repository')

    # Extract zip in working directory
    extract(zip_filepath, work_dir)

    # Commit changes, if any
    if version_changes or has_repos_changes:
        message = create_commit_message(version_changes)
        logger.info('Message: %s', message)

        logger.info('Commit change.')
        commit(work_dir, message, tag, do_commit=not no_commit, push=not no_push)

    # Clean up
    teardown(work_dir, user_work_dir, zip_filepath, user_zip)


if __name__ == '__main__':
    main()
