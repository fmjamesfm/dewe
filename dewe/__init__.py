import sys

try:
    __DEWE_SETUP__
except NameError:
    __DEWE_SETUP__ = False

if __DEWE_SETUP__:
    sys.stderr.write('Running from source directory. \n')
else:
    from .version import git_revision as __git_revision__
    from .version import version as __version__

    from .dewe import * 
    