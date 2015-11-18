from setuptools import setup

setup(
    name='HtmlToWord',
    version='0.6',
    packages=['HtmlToWord', 'HtmlToWord.elements'],
    url='https://github.com/orf/HtmlToWord',
    license='',
    author='Tom',
    author_email='tom@tomforb.es',
    description='Render HTML to a specific portion of a word document',
    install_requires=["BeautifulSoup4",
                      "pypiwin32",
                      "requests"],
    include_package_data=True,
    long_description="""\
Render HTML to a word document using win32com.
Check out the github repo for more information and code samples.
"""
)
