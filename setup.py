from setuptools import setup
from pathlib import Path

this_directory = Path(__file__).parent
long_description = (this_directory / "README.rst").read_text('utf-8')

setup(
    name='frequency_analysis',
    version='0.1.4.5',
    description='Symbol/symbol bigram/word/word bigram frequency analyzer with excel output.',
    long_description=long_description,
    long_description_content_type='text/x-rst',
    author='uqqu',
    classifiers=[
        'Programming Language :: Python :: 3.8',
        'Topic :: Text Processing :: Linguistic',
    ],
    keywords='frequency analysis bigram linguistic cryptanalysis',
    packages=['frequency_analysis'],
    install_requires=['xlsxwriter'],
    url='https://github.com/uqqu/frequency_analysis',
)
