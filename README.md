# Why this code is usefull?
1. Access to all the mails that contains the lead information
2.The data is contained into a gmail direction of asegura2.com, this mail is structured with HTML. For recolect all the data that i need of the lead, i used web scraping,
  with the library bs4 - BeautifulSoup
3. Sent an automatic mail to the lead to start the contact with (used smtplib library to do this)
4. sent an automatic whatsapp message, with the name of the client (with whatsapp api) 
5. Structure all the data with a pandas dataframe
6. And save all the lead data into a excel sheet
