# Automatic Word Document Generator tool python

# :point_right:How to use it?

step 1: Run this code in your terminal:computer:

    git clone https://github.com/Abhimaru/docgeneratortool.git
    cd docgeneratortool
    mkdir LABS,INDIVIDUAL/outputs
    pip install -r requirements.txt
    
step 2: There are two folders we created **LABS:open_file_folder:** and **INDIVIDUAL:open_file_folder:** 
### :point_right: _if you want to create document for the multiple files in same folder then follow this steps:_

1. Inside **"LABS:open_file_folder:"** folder create sub-folders  for example **Lab-1:open_file_folder:**, **Lab-2:open_file_folder:**, **etc.** or paste the folder which having multiple files.
> **Note:closed_book::**
> 1. in **aims**.py:clipboard: file the key value must be your folder name.
> 2. Name of this folders will be heading of the labs in the document.

2. inside that folders paste the files you want to add into word file. and if your lab have screenshots of output then paste into the **output:open_file_folder:** folder.

3. step 4: In *aims*.py:clipboard: file enter the key value as your folder name for example **Lab-1**. Enter value as your aim text

    For Example:
    ["your lab folder name": your aim text"]


### :point_right:_if you want to generate document for individual files then follow this steps:_
1. Inside **"INDIVIDUAL:open_file_folder:"** folder paste the files you want to add in document. follow this naming convention **1-fileName:clipboard:, 2-fileName:clipboard:, 3-fileName:clipboard:, etc...**
> **Note:closed_book::**
> 1. If your Lab has Lab a and b then write foldername as **Lab-No-a**, **Lab-No-b**, etc...

3. inside **OUTPUTS:open_file_folder:** folders paste the output images named as 1.PNG, 2.PNG, 3.PNG and so on. (PNG or JPG) 
> **Note:closed_book::**
If you want to have multiple output image in for one file then give name as **For Example: 2a.PNG, 2b.PNG, etc...**

4.  in **aims**.py:clipboard: enter the key value as **EXPERMENT-1 for file-1, EXPERIMENT-2 for file-2 and so on** .


    For Example:
    ["EXPERIMENT-1": your aim text"]

After that run this code in terminal:computer::

    python generateDoc.py
