# Grad Ball Receipt Automator

## How to Use

1. Simply download the zip file and open it!

  






2. Setting Up

- Navigate to the folder and edit the text file```names.txt``` and replace the sample names it with the names/attendees of those people who you want to generate a Grad Ball Receipt for. Edit ```class_section.txt```, replacing it with your actual words.


- Next, open the file ```Final_GradBall Layout.docx``` in a word processing software such as Microsoft Word and duplicate the content, adding a new page for every attendee. For example, if you have 50 attendees, then duplicate the content by copy and pasting until you have 50 pages. This will probably only take a maximum of 3-4 minutes (simply just Ctrl + C and Ctrl + V)
*Make sure you save the file.*





> [!NOTE]  
> If you only want to review English words, just go to the ```words_cn.txt``` file and DELETE all the words there.





3. Open the Terminal application on your Mac and run the program!



- Install python-docx library

- Navigate to the downloaded folder

```
cd /directorywherethefolderyoudownloadedis
```

For me, it's 
```
cd /users/david/Downloads/spelling-bee-mac-master
```

- Then, run the program using the command 
```
python3 automate.py
```




- If the program finished running successfully, you'll see a document in your folder called ```Modified_Document.docx```.

 Open it and you'll see each attendee with 2 Grad Ball Receipts with a specific class section and 


### And that's it! Happy reviewing! Hope this project inspired you or helped in you in some sort of way!
