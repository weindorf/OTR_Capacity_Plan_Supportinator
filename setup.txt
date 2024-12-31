from setuptools import setup, find_packages

setup(
    name="otr_supportinator",
    version="0.1",
    packages=find_packages(),
    install_requires=[
        'PyQt6',
        # add other dependencies
    ],
    entry_points={
        'console_scripts': [
            'otr_supportinator=otr_supportinator.main:main',
        ],
    },
)
