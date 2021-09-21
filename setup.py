﻿from setuptools import setup

setup(
    name='frequency_analyzer',
    version='0.1',
    description='Symbol/symbol bigram/word/word bigram frequency analyzer with excel output.',
    classifiers=[
        'Programming Language :: Python :: 3.8',
        'Topic :: Text Processing :: Linguistic',
        ],
    keywords='frequancy analysis bigram',
    packages=['frequency_analyzer'],
    install_requires=['xlsxwriter'],
    url='https://github.com/uqqu/frequency_analyzer',
    )