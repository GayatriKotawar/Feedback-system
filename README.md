  Python Feedback System

Overview

-This project is a Feedback Collection System built using Python and Tkinter.
-It allows participants of a Python Workshop to submit their feedback interactively through an emoji-based rating system.
-The feedback responses are saved automatically to an Excel file (Book2.xlsx) using the OpenPyXL library.
-At the end of the survey, users can view a summary of their responses and the overall sentiment (Happy, Neutral, or Unhappy).

Features

-Simple Graphical User Interface (GUI) built with Tkinter
-Emoji-based survey with 10 questions
-Automatic Excel storage for feedback data
-Timestamp added for every feedback submission
-Summary report showing the count of happy, neutral, and sad responses
-Option to restart the survey after submission

Technologies Used

-Python 3.x
-Tkinter – for GUI design
-OpenPyXL – for Excel file handling
-Datetime – to store submission time
-OS module – to check and create Excel file

File Structure

Feedback-System
│
├── feedback_system.py      # Main Python script
└── Book2.xlsx              # Excel file (auto-created if not present)

How It Works

-When the program starts, it checks if Book2.xlsx exists.
        If not, it creates one with proper column headers.
-The user clicks “Start Survey” to begin.
-Each of the 10 questions can be answered using emojis:
     😀 → Happy
     😑 → Neutral
     😔 → Sad
-Once all responses are filled, clicking Submit:
      Calculates the overall sentiment
      Saves the responses with a timestamp into the Excel file
      Displays a summary window with results and an option to restart


The system compares the counts of each emoji and determines:
If 😀 count is highest → Overall Sentiment: Happy
If 😑 count is highest → Overall Sentiment: Neutral
Otherwise → Overall Sentiment: Unhappy


--Excel Output Example

Timestamp	Q1	Q2	Q3	Q4	Q5	Q6	Q7	Q8	Q9	Q10	Summary
2025-10-29 20:30:15	😀	😑	😀	😔	😀	😀	😀	😑	😀	😀	Happy

Author
Gayatri Kotawar
