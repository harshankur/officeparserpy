#!/usr/bin/env python3

"""
Setup script for officeparserpy.

This script uses setuptools to configure the installation and distribution
of the officeparserpy library.
"""

from setuptools import setup

setup(
    name='officeparserpy',
    version='1.0.7',
    author='Harsh Ankur',
    author_email='harshankur@outlook.com',
    description='A Python library to parse text out of any office file. Currently supports docx, pptx, xlsx, odt, odp, ods, pdf files.',
    long_description=open('README.md', 'r').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/harshankur/officeparserpy',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    packages=["officeparserpy"],
    install_requires=[
        'typing-extensions',
        'filetype',
        'pdfminer.six'
    ],
    entry_points={
        'console_scripts': [
            'officeparser = officeparser:main',
        ],
    },
)
