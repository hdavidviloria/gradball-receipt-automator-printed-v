# Grad Ball Receipt Automator

## How to Use

### 1. Simply download the zip file and open it!

  
<img width="744" alt="Screenshot 2024-04-21 at 9 40 38 PM" src="https://github.com/hdavidviloria/gradball-receipt-automator/assets/94162758/78af6694-66d1-4ab2-9b05-cd837c5bc075">

<br></br>





### 2. Setting Up

- Navigate to the folder and edit the text file```names.txt``` and replace the sample names, copy pasting a list you have of the names/attendees of those people who you want to generate a Grad Ball Receipt for. Edit ```class_section.txt```, replacing it with the section of each attendee listed in the names.txt file (in order).

<img width="1236" alt="Screenshot 2024-04-21 at 10 29 41 PM" src="https://github.com/hdavidviloria/gradball-receipt-automator/assets/94162758/2475354e-4ccd-4e28-ac06-2f62acf90f92">


- Next, open the file ```Final_GradBall Layout.docx``` in a word processing software such as Microsoft Word and duplicate the content, adding a new page for every attendee. For example, if you have 50 attendees, then duplicate the content by copy and pasting until you have 50 pages. This will probably only take a maximum of 3-4 minutes (simply just Ctrl + C and Ctrl + V) / Or even less than 30 seconds, if you're really fast at copy pasting just like me :)



https://github.com/hdavidviloria/gradball-receipt-automator/assets/94162758/36088f49-3bbe-4545-86f3-7db1d1da5630




***Make sure you save the file.***





> [!NOTE]  
> You have to ensure that your amount of attendees that you listed in the names.txt file is equal to the number of pages. For example, if a total of 50 attendees so there has to be 50 pages. If there are more pages with content than the number of attendees, Python will give out an error.


<br></br>

### 3. Open the Terminal application on your Mac and run the program!



- Install python-docx library

```
pip3 install python-docx
```



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

 Open it and you'll see each attendee with 2 Grad Ball Receipts with a specific class section and Receipt ID



https://github.com/hdavidviloria/gradball-receipt-automator/assets/94162758/dd4c5eb0-0faa-4071-a08c-f382083b7cf3



### And that's it! Hope this project inspired you or helped in you in some sort of way!
