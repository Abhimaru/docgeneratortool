# Automatic Word Document Generator tool python

# How to use it?

step 1: Run this code in your terminal

    git clone https://github.com/Abhimaru/docgeneratortool.git
    cd docgeneratortool
    mkdir LABS
    pip install -r requirements.txt
    
step 2: Inside **"LABS"** folder create sub-folders for example **Lab-1**, **Lab-2**, **etc.** 
> **Note:**
> 1. in **aims**.py file the key value must be your folder name.
> 2. Name of this folders will be heading of the labs in the document.

step 3: inside that folders paste the files you want to add into word file. and if your lab have screenshots of output then paste into the **output** folder.

step 4: In *aims*.py file enter the key value as your folder name for example **Lab-1**. Enter value as your aim text

    For Example:
    ["your lab folder name": your aim text"]

After that run this code:

    python finalCode.py
