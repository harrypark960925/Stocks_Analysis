# Stock analysis with excel VBA 

## Overview of Project

The purpose of this project was to collect certain stock information in the year of 2017 and 2018 and to make a decision of whether the stock is worth investing or not by using a code of Microsoft Excel VBA. The process was orignially completed in a way without array, however, the goal for this project was to apply the index and arrays to increase the efficiency of the original code. 

## Analysis

For refactored code, I copied the code that was necessary to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps are listed with comments for the structure of refactorization. The images below are the code that I have created. 

<img width="861" alt="Screen Shot 2022-03-07 at 12 13 13 AM" src="https://user-images.githubusercontent.com/83077836/156972072-9f8fac9c-d271-43a4-a940-92b95e294891.png">
<img width="781" alt="Screen Shot 2022-03-07 at 12 17 12 AM" src="https://user-images.githubusercontent.com/83077836/156972415-519bc87d-df23-4ab3-afa9-eb0d9d61b9c1.png">

## Summary 

### Advantages and Disadvantages of refactoring code 

Refacoring the code makes our code much simpler and more organized. Advantages of simpler and organized code are the desgin and debugging, and faster programming. Also, it benefit other users while they view our projects and read the code since it is easier to read with concise information which is very straightforward. However, refactoring the code doesn't always lead to good ways. It is risky when the application is too big and the original existing code does not have the proper test cases. Also, if the developer does not understand the code fully, refactoring the code will lead to a high risk to ruin the program. 

### Advantages of refactoring stocks - analysis

The biggest advantage after refacotoring was the significant decrease of macro running time. The original analysis took approximately 0.5 seconds to run while the new analysis took only about 1/5 time - approximately 0.11 seconds - to run. I've attached the screenshots of the running time for the original analysis and our new analysis. 

<img width="259" alt="Screen Shot 2022-03-07 at 12 27 45 AM" src="https://user-images.githubusercontent.com/83077836/156973606-1c9e286a-c8f4-4c33-b503-907c29932296.png">
<img width="255" alt="Screen Shot 2022-03-07 at 12 28 07 AM" src="https://user-images.githubusercontent.com/83077836/156973615-4b91cc1f-ba6c-4655-bbc5-42668c1e4224.png">
