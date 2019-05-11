import os.path
from setuptools import setup

#The directory containing this file
HERE = os.path.abspath(os.path.dirname(__file__))

# The text of the README file
with open(os.path.join(HERE, "README.md")) as fid:
    README = fid.read()

setup(
   name='ODSForm',
   version='0.1',
   description='Fill out ODS spreadsheets from JSON or YAML',
   license="GPLv3",
   long_description=README,
   long_description_content_type="text/markdown",
   author='Tom Fredrik Blenning',
   author_email='bfg@blenning.no',
   url="https://github.com/BFGConsult/ODSForm",
   packages=['ODSForm'],  #same as name
#   packages=setuptools.find_packages(),
   platforms=['any'],
   install_requires=['pyexcel-ezodf', 'lxml'], #external packages as dependencies
   classifiers=[
       "Programming Language :: Python",
       "Programming Language :: Python :: 2.7",
       "Programming Language :: Python :: 3",
       "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
       "Operating System :: OS Independent",
   ],
#   scripts=[
#            'scripts/cool',
#            'scripts/skype',
#           ]
)
