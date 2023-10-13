# VBA-challenge
#The code is located in the main VBA-challenge repository in the VBA file. 
#Screenshots of the output have also been added. 
#I worked on specific parts of this code with my study group, in particular the For loop that went through each worksheet as well as the Range.value expressions that created the header for my output table. 
#Originally I did not use Doubles for my variables. I changed this after Googling and finding size limits, maybe on stack overflow? 
#The LastRow code to find the last row I took from another in class assignment that Dallin had showed us. I wouldn't have known how to do that otherwise. 
#I also had to Google the formatting of the numbers as a percentage. I believe I originally was multiplying by 100 and then trying to use the .NumberFormat function and that was giving me inaccurate results. I don't recall the exact site/form that I found the answer on but I played with it until I got it to work. 
#Similarly, I had found a forum post, I don't have a direct link to it anymore, that talked about how to use the cells.NumberFormat. If we learned that in class, I hadn't remembered it. 
#For formatting of cells, I referenced this link for formatting code: https://www.automateexcel.com/vba/conditional-formatting/. 
#When preparing my VBA code, I had some trouble if the first line of the data were the largest or smallest value when trying to determine the largest or smallest in the set. After talking to my study group, I see that there are simpler solutions but I think mine still provides the correct answer.
#I did use both the afterhours and the AskBCS. I didn't take anything specifically from the afterhours but AskBCS helped me troubleshoot my code. Specifically, they helped me with the volume variable and when I reset it within the loop. I was confused on why the value was wrong but they helped me move the location of where I reset the value in order to ensure my code was producing correct results. 
