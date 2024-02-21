from setuptools import setup, find_packages

setup(
    name='excel_modifier',
    version='0.1',
    packages=find_packages(),
    description='A package for modifying Excel .xlsm files',
    author='DevNotion',
    author_email='devnotion.info@gmail.com',
    install_requires=[
        'beautifulsoup4>=4.9.3',
    ],
    python_requires='>=3.6',
    include_package_data=True,
    zip_safe=False,
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
    ],
)
