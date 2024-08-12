# ACCURATE-PERCEPTUAL-QUALITY-PREDICTION-FOR-VIDEO-STREAMING-APPLICATIONS
### Overview

Over the course of the year, I developed a Python-based platform using the Tkinter library to collect data for calculating Mean Opinion Scores (MOS) for a diverse video dataset. 

### Key Features

- **Two Interfaces:**
  - **Starter Screen Interface:**
    - Collects user information: Full Name, Age, Class Year, and Race.
    - Utilizes Tkinter for displaying questions and providing answer boxes.
    - Assigns a unique ID to each user, linking their data to their corresponding Excel file.
  - **Main Interface:**
    - Displays videos for user viewing and rating on a scale from 1 to 9.
    - Includes replay functionality for videos.
    - Uses combo boxes for parameter ratings, ensuring that all questions are answered before proceeding.

### Data Management

- **Excel File Storage:**
  - Basic user information is stored in a separate Excel file for each user.
  - The assigned ID ensures correct linkage of data to each user's file.

### User Interaction

- **Video Rating:**
  - Users can watch and rate videos, with parameters presented as combo boxes.
  - Users cannot proceed to the next video until all questions are answered and the video has finished playing.

