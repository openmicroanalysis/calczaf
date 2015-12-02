"""
Script to synchronize CalcZAF source code from Probe Software with GitHub using
Travis CI.
Steps:

  * Pull repository from GitHub
  * Download ZIP containing CalcZAF source code from Probe Software website
  * Parse version text file and compare for new changes
  * Add modified files, commit and push to GitHub
"""

# Standard library modules.
import logging
logger = logging.getLogger(__name__)
import os
import sys
import urllib.request
import urllib.parse
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

        if change_pattern.match(line):
            lines = []
            dt, line = line.split('\t', 1)
            try:
                dt = datetime.datetime.strptime(dt, '%m/%d/%y')
            except:
                dt = None
                continue
            changes[dt] = lines
            lines.append(line)
        elif tag_pattern.match(line):
            if dt is not None:
                tags[dt] = line[2:].split('\t', 1)[0].strip()
        else:
            lines.append(line)

    return changes, tags

def pull(workdir, reposurl):
    logger.info('Running git pull/clone...')
    if '.git' in os.listdir(workdir):
        args = ['git', 'pull']
    else:
        args = ['git', 'clone', reposurl, workdir]
    subprocess.check_call(args, cwd=workdir)
    logger.info('Running git pull/clone... DONE')

def commit(workdir, message, tag, push=True):
    logger.info('Running git add...')
    args = ['git', 'add', '.']
    subprocess.check_call(args, cwd=workdir)
    logger.info('Running git add... DONE')

    logger.info('Running git commit...')
    args = ['git', 'commit', '-m', message]
    subprocess.check_call(args, cwd=workdir)
    logger.info('Running git commit... DONE')

    if tag is not None:
        logger.info('Running git tag...')
        args = ['git', 'tag', tag]
        subprocess.check_call(args, cwd=workdir)
        logger.info('Running git tag... DONE')

    if push:
        logger.info('Running git push...')
        args = ['git', 'push', '--all']
        subprocess.check_call(args, cwd=workdir)
        logger.info('Running git push... DONE')

def compare(filepath, workdir):
    logger.info('Reading current version...')
    workfilepath = os.path.join(workdir, VERSION_FILENAME)
    with open(workfilepath, 'r', errors='ignore') as fp:
        oldchanges, oldtags = parse_version(fp)

    latestchange = max(oldchanges)
    latesttag = max(oldtags)
    logger.info('Reading current version... DONE')

    logger.info('Reading zip version...')
    with zipfile.ZipFile(filepath, 'r') as z:
        with io.TextIOWrapper(z.open(VERSION_FILENAME), errors='ignore') as fp:
            newchanges, newtags = parse_version(fp)
    logger.info('Reading zip version... DONE')

    logger.info('Comparing versions...')
    newchanges = dict((dt, val) for dt, val in newchanges.items()
                      if dt > latestchange)
    newtags = dict((dt, val) for dt, val in newtags.items()
                   if dt > latesttag)
    newtag = newtags[max(newtags)] if newtags else None
    logger.info('Comparing versions... DONE')

    return newchanges, newtag

def extract(filepath, workdir):
    logger.info('Extracting %s...', filepath)
    with zipfile.ZipFile(filepath, 'r') as z:
        z.extractall(workdir)
    logger.info('Extracting %s... DONE', filepath)

def all_files_lowercase(workdir):
    logger.info("all_files_lowercase")

    ignore_paths = [".git", ".travis"]
    ignore_files = [".gitignore", ".travis.yml", "VERSION.TXT"]
    rename_file_extentions = ['.bas', '.frm', '.frx', '.vbp', '*.vbw']

    for root, dirs, files in os.walk(workdir):
        for ignore_path in ignore_paths:
            if ignore_path in root:
                break

            for filename in files:
                if filename not in ignore_files:
                    _root, extention = os.path.splitext(filename)
                    if extention.lower() in rename_file_extentions and filename != filename.lower():
                        srcpath = os.path.join(root, filename)
                        dstpath = os.path.join(root, filename.lower())
                        if os.path.isfile(srcpath):
                            logger.info("Renaming %s to %s", filename, filename.lower())
                            os.rename(srcpath, dstpath)

def create_commit_message(changes):
    logger.info('Create commit message...')

    message = []
    message += ['Auto-sync on %s' % \
                    datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S'),
                '']

    for dt in sorted(changes):
        message.append(dt.strftime('%Y/%m/%d'))
        for line in changes[dt]:
            message.append(' ' * 4 + line)

    logger.info('Create commit message... DONE')

    return '\n'.join(message)

def download(url):
    def reporthook(blocknum, blocksize, totalsize):
        """
        From http://stackoverflow.com/questions/13881092/download-progressbar-for-python-3
        """
        readsofar = blocknum * blocksize
        if totalsize > 0:
            percent = readsofar * 1e2 / totalsize
            s = "\r%5.1f%% %*d / %d" % (
                percent, len(str(totalsize)), readsofar, totalsize)
            sys.stderr.write(s)
            if readsofar >= totalsize: # near the end
                sys.stderr.write("\n")
        else: # total size is unknown
            sys.stderr.write("read %d\n" % (readsofar,))

    logger.info('Downloading...')
    filepath, _headers = urllib.request.urlretrieve(url, reporthook=reporthook)
    logger.info('Downloading... DONE')

    return filepath

def setup(workdir):
    userworkdir = True
    if not workdir:
        workdir = tempfile.mkdtemp()
        userworkdir = False
        logger.info('Created work directory: %s', workdir)

    if not os.path.exists(workdir):
        os.makedirs(workdir)
        logger.info('Created work directory: %s', workdir)

    return workdir, userworkdir

def teardown(workdir, userworkdir, zipfilepath, userzip):
    if not userworkdir:
        shutil.rmtree(workdir)
        logger.info('Removed work directory: %s', workdir)
    if not userzip:
        os.remove(zipfilepath)

def main():
    description = 'Sync CalcZAF from Probe Software to openMicroanalysis GitHub'
    parser = argparse.ArgumentParser(description=description)

    parser.add_argument('-v', '--verbose', action='store_true', default=False,
                        help='Verbose mode')
    parser.add_argument('-w', '--workdir', help='Working directory')
    parser.add_argument('-u', '--url',
                        default='http://probesoftware.com/download/CALCZAF_SOURCE-E2.ZIP',
                        help='Url to CalcZAF source zip file')
    parser.add_argument('--reposurl', default='git@github.com:openmicroanalysis/calczaf.git')
    parser.add_argument('--no-push', action='store_true', default=False,
                        help='Do not push after commit')
    parser.add_argument('--no-commit', action='store_true', default=False,
                        help='Do not commit change')
    parser.add_argument('--force-check-change', action='store_true', default=False,
                        help='Force checking change even if the version is the same')

    args = parser.parse_args()


    level = logging.DEBUG if args.verbose else logging.INFO
    logger.addHandler(logging.StreamHandler())
    logger.setLevel(level)

    no_commit = args.no_commit
    force_check_change = args.force_check_change
    logger.info(no_commit)
    logger.info(force_check_change)

    workdir = args.workdir
    workdir, userworkdir = setup(workdir)

    # Pull repository
    reposurl = args.reposurl
    pull(workdir, reposurl)

    # Download zip
    url = args.url
    userzip = True
    if urllib.parse.urlparse(url).scheme == 'd':
        zipfilepath = url
    elif urllib.parse.urlparse(url).scheme != '':
        zipfilepath = download(url)
        userzip = False
    else:
        zipfilepath = url

    # Compare versions
    changes, tag = compare(zipfilepath, workdir)
    if not changes and not force_check_change:
        teardown(workdir, userworkdir, zipfilepath, userzip)
        logger.info('No change. Exiting.')
        return

    if changes or force_check_change:
        logger.info('%i new changes found', len(changes))

        extract(zipfilepath, workdir)

        all_files_lowercase(workdir)

        message = create_commit_message(changes)
        logger.info('Message: %s', message)

        push = not args.no_push
        if not no_commit:
            logger.info('Commit change.')
            commit(workdir, message, tag, push=push)

    # Clean up
    teardown(workdir, userworkdir, zipfilepath, userzip)

if __name__ == '__main__':
    main()
