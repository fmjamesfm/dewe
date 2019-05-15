import subprocess
import os
import builtins

PACKAGE_DIR         = 'dewe'
PACKAGE_NAME        = 'dewe'
MAJOR               = 0
MINOR               = 1
MICRO               = 0
VERSION             = '%d.%d.%d' % (MAJOR, MINOR, MICRO)

__author__ = 'James Massaglia'
__email__ = 'fakeitalian@carbonair.eu'
__url__ = 'https://gitlab.com/carbonair/dewe'

# make sure cutils doesnt import all modules during setup
builtins.__DEWE_SETUP__ = True

# Return the git revision as a string
def git_version():
    def _minimal_ext_cmd(cmd):
        # construct minimal environment
        env = {}
        for k in ['SYSTEMROOT', 'PATH', 'HOME']:
            v = os.environ.get(k)
            if v is not None:
                env[k] = v
        # LANGUAGE is used on win32
        env['LANGUAGE'] = 'C'
        env['LANG'] = 'C'
        env['LC_ALL'] = 'C'
        pr = subprocess.Popen(cmd, stdout=subprocess.PIPE, env=env)
        return (pr.communicate()[0], pr.returncode)

    try:
        out, _ = _minimal_ext_cmd(['git', 'rev-parse', 'HEAD'])
        GIT_REVISION = out.strip().decode('ascii')
        if _minimal_ext_cmd(['git', 'diff-files', '--quiet'])[-1]:
            GIT_REVISION += '.dirty'
    except OSError:
        GIT_REVISION = "Unknown"

    return GIT_REVISION

def get_version_info():
    # Adding the git rev number needs to be done inside write_version_py(),
    # otherwise the import of numpy.version messes up the build under Python 3.
    FULLVERSION = VERSION
    if os.path.exists('.git'):
        GIT_REVISION = git_version()
    else:
        GIT_REVISION = "Unknown"

    return FULLVERSION, GIT_REVISION

def write_version_py(filename=PACKAGE_DIR+'/version.py'):
    cnt = """
# THIS FILE IS GENERATED FROM {} SETUP.PY

short_version = '%(version)s'
version = '%(version)s'
full_version = '%(full_version)s'
git_revision = '%(git_revision)s'
""".format(PACKAGE_NAME.upper())
    FULLVERSION, GIT_REVISION = get_version_info()

    a = open(filename, 'w')
    try:
        a.write(cnt % {'version': VERSION,
                       'full_version': FULLVERSION,
                       'git_revision': GIT_REVISION})
    finally:
        a.close()

def setup_package():
	write_version_py()

	metadata = dict(
		name=PACKAGE_NAME,
		version=get_version_info()[0],
		description='Dewesoft COM control module',
		url=__url__,
		author=__author__,
		author_email=__email__,
		license='',
		packages=[PACKAGE_DIR],
		install_requires=[
		'pandas',
        'numpy',
        'scipy',
        'sounddevice',
        'pywin32',
        'pint'
		  ],
          package_data = {PACKAGE_DIR: ['mod_defs_en.txt', 'constants_en.txt']},
		  zip_safe=False
	)

	from setuptools import setup
	metadata['version'] = get_version_info()[0]

	setup(**metadata)


if __name__ == '__main__':
	setup_package()
	del builtins.__DEWE_SETUP__
