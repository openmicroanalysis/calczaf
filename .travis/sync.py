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

def has_changes(workdir):
    logger.info('Running git status...')
    args = ['git', 'status', '--porcelain']
    output = subprocess.check_output(args, cwd=workdir, universal_newlines=True)
    logger.info('Running git status... DONE')
    return bool(output)

def commit(workdir, message, tag, commit=True, push=True):
    logger.info('Running git add...')
    args = ['git', 'add', '.']
    subprocess.check_call(args, cwd=workdir)
    logger.info('Running git add... DONE')

    logger.info('Running git commit...')
    args = ['git', 'commit', '-m', message]
    if not commit: args.append('--dry-run')
    subprocess.check_call(args, cwd=workdir)
    logger.info('Running git commit... DONE')

    if tag is not None and commit:
        logger.info('Running git tag...')
        args = ['git', 'tag', tag]
        subprocess.check_call(args, cwd=workdir)
        logger.info('Running git tag... DONE')

    logger.info('Running git push...')
    args = ['git', 'push', '--all']
    if not push: args.append('--dry-run')
    subprocess.check_call(args, cwd=workdir)
    logger.info('Running git push... DONE')

def compare(filepath, workdir):
    logger.info('Reading current version...')
    workfilepath = os.path.join(workdir, VERSION_FILENAME.lower())
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

def compare_remove_files(filepath, workdir, no_commit):
    logger.info('Comparing files %s...', filepath)

    zipfile_list = set()
    with zipfile.ZipFile(filepath, 'r') as z:
        for info in z.infolist():
            # Always use lower case
            filename = info.filename.lower()
            zipfile_list.add(filename)
            # Create directory if needed
            dirname = os.path.join(workdir, os.path.dirname(filename))
            os.makedirs(dirname, exist_ok=True)

    workdir_list = set(os.listdir(workdir))
    # Ignore these files
    workdir_list.remove(".travis.yml")
    workdir_list.remove(".travis")
    workdir_list.remove(".git")
    workdir_list.remove(".gitignore")
    workdir_list.remove("readme.md")
    workdir_list.remove("license")

    removed_files = workdir_list - zipfile_list

    logger.info('Comparing files %s... DONE', filepath)

    if len(removed_files) > 0:
        for removed_file in removed_files:
            logger.info('Running git rm %s ...', removed_file)
            args = ['git', 'rm', removed_file]
            if no_commit: args.append('--dry-run')
            subprocess.check_call(args, cwd=workdir)
            logger.info('Running git rm %s... DONE', removed_file)

        if not no_commit:
            message = "Remove files not in the CalcZAF source zip file."
            logger.info('Running git commit...')
            args = ['git', 'commit', '-m', message]
            subprocess.check_call(args, cwd=workdir)
            logger.info('Running git commit... DONE')

def extract(filepath, workdir):
    logger.info('Extracting %s...', filepath)
    with zipfile.ZipFile(filepath, 'r') as z:
        for info in z.infolist():
            # Always use lower case
            filename = info.filename.lower()

            # Create directory if needed
            dirname = os.path.join(workdir, os.path.dirname(filename))
            os.makedirs(dirname, exist_ok=True)

            # Write filename
            dstpath = os.path.join(workdir, filename)
            with z.open(info, 'r') as fi, open(dstpath, 'wb') as fo:
                fo.write(fi.read())

    logger.info('Extracting %s... DONE', filepath)

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
    if no_pull: logger.info('Pull disabled')
    no_commit = args.no_commit
    if no_commit: logger.info('Commit disabled')
    no_push = args.no_push
    if no_push: logger.info('Push disabled')

    workdir = args.workdir
    workdir, userworkdir = setup(workdir)

    # Pull repository
    reposurl = args.reposurl
    if not no_pull:
        pull(workdir, reposurl)

    # Download zip
    url = args.url
    try:
        zipfilepath = download(url)
        userzip = False
    except urllib.error.URLError:
        zipfilepath = url
        userzip = True
    except ValueError:
        zipfilepath = url
        userzip = True

    # Compare versions
    version_changes, tag = compare(zipfilepath, workdir)
    logger.info('%i new changes found in VERSION.TXT', len(version_changes))
    
    # Compare files
    compare_remove_files(zipfilepath, workdir, no_commit)

    has_repos_changes = has_changes(workdir)
    if has_repos_changes:
        logger.info('has changes found in repository')

    # Extract zip in working directory
    extract(zipfilepath, workdir)

    # Commit changes, if any
    if version_changes or has_repos_changes:
        message = create_commit_message(version_changes)
        logger.info('Message: %s', message)

        logger.info('Commit change.')
        commit(workdir, message, tag, commit=not no_commit, push=not no_push)

    # Clean up
    teardown(workdir, userworkdir, zipfilepath, userzip)

if __name__ == '__main__':
    main()
