import setuptools

# with open("README.md", "r",encoding="utf8") as fh:
#     long_description = fh.read()

setuptools.setup(
    name="speech4t",
    version="0.0.1",
    author="Wen-Hung, Chang 張文宏",
    author_email="beardad1975@nmes.tyc.edu.tw",
    description="Speech wrapper for Teenagers",
    long_description="Speech wrapper for Teenagers",
    long_description_content_type="text/markdown",
    url="https://github.com/beardad1975/speech4t",
    #packages=setuptools.find_packages(),
    platforms=["Windows"],
    python_requires=">=3.5",
    packages=['speech4t','語音模組'],
    install_requires = ['pywin32>=228', 'SpeechRecognition>=3.8.1'],
        classifiers=[
            "Programming Language :: Python :: 3",
            "License :: OSI Approved :: MIT License",
            "Operating System :: Microsoft :: Windows",
            #"Operating System :: MacOS",
            #"Operating System :: POSIX :: Linux",
            "Natural Language :: Chinese (Traditional)",
        ],
)